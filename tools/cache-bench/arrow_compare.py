# Real Arrow and Parquet on made-up data, entirely in memory (nothing is written to disk).
# Uses the pyarrow already on this PC; installs nothing. Same data rules as com-bench.
#   python -I arrow_compare.py <rows>
import ctypes, random, statistics, sys, time
import pyarrow as pa
import pyarrow.compute as pc
import pyarrow.ipc as ipc
import pyarrow.parquet as pq
from ctypes import wintypes as wt


class MS(ctypes.Structure):
    _fields_ = [("len", wt.DWORD), ("load", wt.DWORD)] + [(n, ctypes.c_ulonglong) for n in "tp ap tf af tv av ae".split()]


class PMC(ctypes.Structure):
    _fields_ = [("cb", wt.DWORD), ("pf", wt.DWORD)] + [(n, ctypes.c_size_t) for n in "peak ws qpp qp qnp qn pu ppu priv".split()]


def commit():
    m = MS(); m.len = ctypes.sizeof(MS); ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
    return 100.0 * (m.tf - m.af) / m.tf


k32 = ctypes.WinDLL("kernel32", use_last_error=True)
k32.GetCurrentProcess.restype = wt.HANDLE
k32.K32GetProcessMemoryInfo.argtypes = [wt.HANDLE, ctypes.POINTER(PMC), wt.DWORD]


def priv_mb():
    p = PMC(); p.cb = ctypes.sizeof(PMC)
    k32.K32GetProcessMemoryInfo(k32.GetCurrentProcess(), ctypes.byref(p), p.cb)
    return p.priv // 1048576


def gate():
    c = commit()
    if c >= 90:
        print("STOP: commit >= 90%"); sys.exit(3)
    n = 0
    while c >= 85:
        if n == 0: print("waiting: commit >= 85%")
        time.sleep(1); n += 1; c = commit()
        if n > 300: print("SKIPPED: commit stayed high"); sys.exit(3)


rows = int(sys.argv[1]); cols = 20
print(f"pyarrow {pa.__version__}; {rows}x{cols}; commit {commit():.1f}%; python+pyarrow {priv_mb()}MB")

# Excel cells are not typed by column. Can Arrow hold a mixed column without extra work?
try:
    pa.array([1.5, "text", None, True]); print("mixed column: accepted")
except Exception as e:
    print(f"mixed column as a plain array: REFUSED ({type(e).__name__})")

t0 = time.perf_counter(); arrays = []; names = []; rng = random.Random(1)
for c in range(cols):
    gate(); k = c % 10
    if k < 5: a = pa.array([round(rng.random() * 100000, 3) for _ in range(rows)], type=pa.float64())
    elif k < 8: a = pa.array([f"GROUP-{rng.randrange(1000)}" for _ in range(rows)]).dictionary_encode()
    else: a = pa.array([f"ITEM-{i + 1}-{c}" for i in range(rows)], type=pa.string())
    arrays.append(a); names.append(f"c{c}")
batches_per = 1000
table = pa.table(arrays, names=names); del arrays
print(f"make table {time.perf_counter() - t0:.1f}s; table data {table.nbytes / 1048576:.0f}MB ({table.nbytes / (rows * cols):.1f} bytes/cell); arrow allocated {pa.total_allocated_bytes() / 1048576:.0f}MB; process {priv_mb()}MB; commit {commit():.1f}%")


def screens(t, label):
    r = random.Random(9); s = []
    for _ in range(500):
        p = r.randrange(0, rows - 40); t0 = time.perf_counter(); t.slice(p, 40).to_pylist(); s.append((time.perf_counter() - t0) * 1000)
    s.sort(); print(f"{label}: screen of 40 rows median={s[250]:.2f}ms p95={s[475]:.2f}ms max={s[-1]:.2f}ms")


screens(table, "table in memory")
t0 = time.perf_counter(); hits = 0
for c in table.columns:
    if pa.types.is_string(c.type): hits += pc.sum(pc.match_substring(c, "NO-SUCH-VALUE-ZZZ")).as_py() or 0
    elif pa.types.is_dictionary(c.type): hits += pc.sum(pc.match_substring(c.combine_chunks().dictionary, "NO-SUCH-VALUE-ZZZ")).as_py() or 0
print(f"search all text, nothing found: {(time.perf_counter() - t0) * 1000:.0f}ms hits={hits}")
t0 = time.perf_counter(); pc.sort_indices(table["c0"]); print(f"sort by a number column: {(time.perf_counter() - t0) * 1000:.0f}ms")
t0 = time.perf_counter(); pc.sort_indices(table["c5"].cast(pa.string())); print(f"sort by a repeated-text column: {(time.perf_counter() - t0) * 1000:.0f}ms")
t0 = time.perf_counter(); pc.sort_indices(table["c8"]); print(f"sort by a unique-text column: {(time.perf_counter() - t0) * 1000:.0f}ms")

# Arrow IPC (Feather V2 is this format) in a memory buffer.
gate(); t0 = time.perf_counter(); sink = pa.BufferOutputStream()
with ipc.new_file(sink, table.schema) as w: w.write_table(table, max_chunksize=1000)
buf = sink.getvalue(); wr = time.perf_counter() - t0
t0 = time.perf_counter(); rd = ipc.open_file(buf); t2 = rd.read_all(); op = time.perf_counter() - t0
print(f"Arrow IPC: write {wr:.2f}s size {buf.size / 1048576:.0f}MB; open+map all {op * 1000:.1f}ms")
t0 = time.perf_counter(); rd2 = ipc.open_file(buf); rd2.get_batch(0).slice(0, 40).to_pylist(); print(f"Arrow IPC: first screen from buffer {(time.perf_counter() - t0) * 1000:.2f}ms")
del t2, rd, rd2, buf, sink

# Parquet + Zstd in a memory buffer.
gate(); t0 = time.perf_counter(); sink = pa.BufferOutputStream(); pq.write_table(table, sink, compression="zstd", row_group_size=100000); buf = sink.getvalue(); wr = time.perf_counter() - t0
del table
t0 = time.perf_counter(); pf = pq.ParquetFile(pa.BufferReader(buf)); first = pf.read_row_group(0).slice(0, 40).to_pylist(); fs = time.perf_counter() - t0
t0 = time.perf_counter(); whole = pq.read_table(pa.BufferReader(buf)); rd = time.perf_counter() - t0
print(f"Parquet+Zstd: write {wr:.2f}s size {buf.size / 1048576:.0f}MB; first screen {fs * 1000:.0f}ms (must unpack one group); read all {rd:.2f}s; memory after {pa.total_allocated_bytes() / 1048576:.0f}MB; process {priv_mb()}MB")
print(f"commit at end {commit():.1f}%")
