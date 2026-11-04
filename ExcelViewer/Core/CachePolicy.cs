// The one place allowed to create, replace or delete a cache file.
//
// THE RULE: a cache holds company data with no NASCA protection. The owner, acting as IT manager,
// approved it on 2026-11-04 for this laptop's local disk ONLY. Nothing may leave this machine.
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

namespace ExcelViewer.Core;

public sealed class CachePolicyException : Exception { public CachePolicyException(string m) : base(m) { } }

public static class CachePolicy
{
    public const string Folder = @"D:\ExcelViewerCache";
    public const string KeptTestCopy = "real-test.xvc"; // the owner asked to keep this one
    const string ApprovedMachineHash = "B2ACEA4AF886F49CB14C97206C200E9578396BFBD2EA05F827A9A694F38F3B26";

    static string Sha(string s) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(s)));

    public static bool ThisPcApproved => Sha(Environment.MachineName.ToUpperInvariant()) == ApprovedMachineHash;

    /// <summary>The cache name is a hash, so the source file's name is not visible on disk.</summary>
    public static string PathFor(string source) =>
        Path.Combine(Folder, Sha(Path.GetFullPath(source).ToUpperInvariant())[..32] + ".xvc");

    /// <summary>Throws unless this path is a file directly in the approved folder, on this laptop's fixed disk.</summary>
    public static void Validate(string path)
    {
        if (!ThisPcApproved) throw new CachePolicyException("The disk copy is approved for one laptop only, and this is not that laptop.");
        string full = Path.GetFullPath(path);
        if (!string.Equals(Path.GetDirectoryName(full), Folder, StringComparison.OrdinalIgnoreCase))
            throw new CachePolicyException("A cache may only be written to the approved local folder.");
        if (!(full.EndsWith(".xvc", StringComparison.OrdinalIgnoreCase) || full.EndsWith(".xvc.part", StringComparison.OrdinalIgnoreCase)))
            throw new CachePolicyException("Only cache files may be written to the approved folder.");
        if (string.Equals(Path.GetFileName(full), KeptTestCopy, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Path.GetFileName(full), KeptTestCopy + ".part", StringComparison.OrdinalIgnoreCase))
            throw new CachePolicyException("The owner's kept test copy must not be changed.");
        if (new DriveInfo(Path.GetPathRoot(Folder)).DriveType != DriveType.Fixed)
            throw new CachePolicyException("The approved folder is not on a fixed local disk.");
    }

    /// <summary>Creates the folder if needed, and makes sure only this account and Windows can read it.</summary>
    public static void EnsureFolderLocked()
    {
        var dir = new DirectoryInfo(Folder); if (!dir.Exists) dir.Create();
        var me = WindowsIdentity.GetCurrent().User; var system = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
        var sec = dir.GetAccessControl();
        var rules = sec.GetAccessRules(true, true, typeof(SecurityIdentifier)).Cast<FileSystemAccessRule>().ToList();
        bool locked = sec.AreAccessRulesProtected && rules.All(r => r.IdentityReference.Equals(me) || r.IdentityReference.Equals(system));
        if (locked) return;
        sec.SetAccessRuleProtection(true, false);
        foreach (var r in rules) sec.RemoveAccessRuleAll(r);
        const InheritanceFlags inh = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;
        sec.AddAccessRule(new FileSystemAccessRule(me, FileSystemRights.FullControl, inh, PropagationFlags.None, AccessControlType.Allow));
        sec.AddAccessRule(new FileSystemAccessRule(system, FileSystemRights.FullControl, inh, PropagationFlags.None, AccessControlType.Allow));
        dir.SetAccessControl(sec);
    }

    public static FileStream OpenForWrite(string path)
    {
        Validate(path); EnsureFolderLocked();
        return new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20);
    }

    public static void ReplaceWith(string partPath, string finalPath)
    {
        Validate(partPath); Validate(finalPath);
        File.Move(partPath, finalPath, true);
    }

    public static void Delete(string path)
    {
        Validate(path);
        if (File.Exists(path)) File.Delete(path);
    }
}
