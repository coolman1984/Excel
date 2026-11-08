using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelApp.Formula;

public enum ValueKind { Blank, Number, Text, Bool, Error, Unsupported }
public readonly record struct FValue(ValueKind Kind, double Number = 0, string? Text = null)
{
    public static FValue Blank => new(ValueKind.Blank); public static FValue Unsupported => new(ValueKind.Unsupported);
    public static FValue Num(double n) => new(ValueKind.Number, n); public static FValue Str(string s) => new(ValueKind.Text, 0, s); public static FValue Bool(bool b) => new(ValueKind.Bool, b ? 1 : 0); public static FValue Error(string e) => new(ValueKind.Error, 0, e);
}
public interface IFormulaGrid { FValue Get(int sheet, int row, int col); int SheetIndex(string name); }
public abstract record Node;
public sealed record Literal(FValue Value) : Node;
public sealed record Reference(string? Sheet, int Row, int Col) : Node;
public sealed record Range(Reference Start, Reference End) : Node;
public sealed record Unary(string Op, Node Value) : Node;
public sealed record Binary(string Op, Node Left, Node Right) : Node;
public sealed record Call(string Name, IReadOnlyList<Node> Args) : Node;
public sealed record Token(string Kind, string Text);
public static class Lexer
{
    public static IReadOnlyList<Token> Lex(string formula)
    {
        var tokens = new List<Token>(); int i = formula.StartsWith('=') ? 1 : 0;
        while (i < formula.Length)
        {
            char c = formula[i]; if (char.IsWhiteSpace(c)) { i++; continue; }
            if (c == '"') { i++; var value = ""; bool closed = false; while (i < formula.Length) { if (formula[i] == '"') { i++; if (i < formula.Length && formula[i] == '"') { value += '"'; i++; continue; } closed = true; break; } value += formula[i++]; } if (!closed) throw new FormatException("Unclosed string."); tokens.Add(new("string", value)); continue; }
            if (c == '\'') { i++; var value = ""; bool closed = false; while (i < formula.Length) { if (formula[i] == '\'') { i++; if (i < formula.Length && formula[i] == '\'') { value += '\''; i++; continue; } closed = true; break; } value += formula[i++]; } if (!closed) throw new FormatException("Unclosed sheet name."); tokens.Add(new("name", value)); continue; }
            if (char.IsDigit(c) || (c == '.' && i + 1 < formula.Length && char.IsDigit(formula[i + 1]))) { var match = Regex.Match(formula[i..], @"^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?"); tokens.Add(new("number", match.Value)); i += match.Length; continue; }
            if (char.IsLetter(c) || c is '_' or '$' or '#') { int start = i++; while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '.' or '$' || (c == '#' && formula[i] is '/' or '?' or '!'))) { if (formula[i] == '!' && c != '#') break; i++; } tokens.Add(new("name", formula[start..i])); continue; }
            string op = i + 1 < formula.Length && (formula.Substring(i, 2) is "<=" or ">=" or "<>") ? formula.Substring(i, 2) : c.ToString(); i += op.Length; if (!"+-*/^&=<>%(),:!".Contains(op[0])) throw new FormatException("Invalid formula character."); tokens.Add(new(op, op));
        }
        tokens.Add(new("end", "")); return tokens;
    }
}
public sealed class Parser
{
    readonly IReadOnlyList<Token> tokens; int pos;
    public Parser(string formula) => tokens = Lexer.Lex(formula);
    Token Current => tokens[pos]; bool Take(string kind) { if (Current.Kind != kind) return false; pos++; return true; }
    void Need(string kind) { if (!Take(kind)) throw new FormatException("Invalid formula."); }
    public Node Parse() { var node = Compare(); Need("end"); return node; }
    Node Compare() { var n = Concat(); while (Current.Kind is "=" or "<>" or "<" or ">" or "<=" or ">=") { var op = tokens[pos++].Kind; n = new Binary(op, n, Concat()); } return n; }
    Node Concat() { var n = Sum(); while (Take("&")) n = new Binary("&", n, Sum()); return n; }
    Node Sum() { var n = Product(); while (Current.Kind is "+" or "-") { string op = tokens[pos++].Kind; n = new Binary(op, n, Product()); } return n; }
    Node Product() { var n = Power(); while (Current.Kind is "*" or "/") { string op = tokens[pos++].Kind; n = new Binary(op, n, Power()); } return n; }
    Node Power() { var n = Prefix(); while (Take("^")) n = new Binary("^", n, Prefix()); return n; }
    Node Prefix() { if (Take("-")) return new Unary("-", Prefix()); if (Take("+")) return Prefix(); var n = Primary(); while (Take("%")) n = new Unary("%", n); return n; }
    static Reference Ref(string? sheet, string text) { var match = Regex.Match(text, @"^\$?([A-Za-z]{1,3})\$?(\d+)$"); if (!match.Success) throw new FormatException("Invalid cell reference."); int col = 0; foreach (char c in match.Groups[1].Value.ToUpperInvariant()) col = col * 26 + c - 'A' + 1; int row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture); if (row < 1 || row > 1048576 || col > 16384) throw new FormatException("Invalid cell reference."); return new Reference(sheet, row - 1, col - 1); }
    Node Primary()
    {
        if (Take("(")) { var n = Compare(); Need(")"); return n; }
        if (Current.Kind == "number") return new Literal(FValue.Num(double.Parse(tokens[pos++].Text, CultureInfo.InvariantCulture)));
        if (Current.Kind == "string") return new Literal(FValue.Str(tokens[pos++].Text));
        if (Current.Kind != "name") throw new FormatException("Invalid formula."); string name = tokens[pos++].Text;
        if (Take("(")) { var args = new List<Node>(); if (!Take(")")) { do { args.Add(Compare()); } while (Take(",")); Need(")"); } return new Call(name.ToUpperInvariant(), args); }
        if (name.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return new Literal(FValue.Bool(true)); if (name.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return new Literal(FValue.Bool(false)); if (name.StartsWith('#')) return new Literal(FValue.Error(name));
        string? sheet = null; if (Take("!")) { sheet = name; if (Current.Kind != "name") throw new FormatException("Invalid reference."); name = tokens[pos++].Text; }
        if (Regex.IsMatch(name, @"^\$?[A-Za-z]{1,3}$") && Take(":")) { string end = tokens[pos++].Text; var a = Ref(sheet, name + "1"); var b = Ref(sheet, end + "1048576"); return new Range(a, b); }
        var first = Ref(sheet, name); if (!Take(":")) return first; if (Current.Kind != "name") throw new FormatException("Invalid range."); return new Range(first, Ref(sheet, tokens[pos++].Text));
    }
}
