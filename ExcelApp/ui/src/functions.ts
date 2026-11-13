// The worksheet functions the editor calculates, with Excel's argument hints. Keep in step with
// ExcelApp/Formula/Evaluator.cs and EvaluatorMore.cs (a test checks the names match the host's list).

export interface FnInfo {
    name: string;
    args: string;
    desc: string;
    cat: "Math" | "Statistical" | "Text" | "Logical" | "Lookup" | "Date & Time" | "Information" | "Financial";
}

const F = (cat: FnInfo["cat"], name: string, args: string, desc: string): FnInfo => ({ name, args, desc, cat });

export const FUNCTIONS: FnInfo[] = [
    // Math & trig
    F("Math", "ABS", "number", "Returns the absolute value of a number."),
    F("Math", "ACOS", "number", "Returns the arccosine of a number, in radians."),
    F("Math", "ASIN", "number", "Returns the arcsine of a number, in radians."),
    F("Math", "ATAN", "number", "Returns the arctangent of a number, in radians."),
    F("Math", "ATAN2", "x_num, y_num", "Returns the arctangent of x and y coordinates, in radians."),
    F("Math", "CEILING", "number, significance", "Rounds a number up to the nearest multiple of significance."),
    F("Math", "CEILING.MATH", "number, [significance], [mode]", "Rounds a number up to the nearest integer or multiple of significance."),
    F("Math", "COS", "number", "Returns the cosine of an angle."),
    F("Math", "DEGREES", "angle", "Converts radians to degrees."),
    F("Math", "EVEN", "number", "Rounds a number up to the nearest even integer."),
    F("Math", "EXP", "number", "Returns e raised to the power of a number."),
    F("Math", "FACT", "number", "Returns the factorial of a number."),
    F("Math", "FLOOR", "number, significance", "Rounds a number down to the nearest multiple of significance."),
    F("Math", "FLOOR.MATH", "number, [significance], [mode]", "Rounds a number down to the nearest integer or multiple of significance."),
    F("Math", "GCD", "number1, [number2], ...", "Returns the greatest common divisor."),
    F("Math", "INT", "number", "Rounds a number down to the nearest integer."),
    F("Math", "LCM", "number1, [number2], ...", "Returns the least common multiple."),
    F("Math", "LN", "number", "Returns the natural logarithm of a number."),
    F("Math", "LOG", "number, [base]", "Returns the logarithm of a number to the base you give."),
    F("Math", "LOG10", "number", "Returns the base-10 logarithm of a number."),
    F("Math", "MOD", "number, divisor", "Returns the remainder after a number is divided by a divisor."),
    F("Math", "MROUND", "number, multiple", "Returns a number rounded to the desired multiple."),
    F("Math", "ODD", "number", "Rounds a number up to the nearest odd integer."),
    F("Math", "PI", "", "Returns the value of pi."),
    F("Math", "POWER", "number, power", "Returns the result of a number raised to a power."),
    F("Math", "PRODUCT", "number1, [number2], ...", "Multiplies all the numbers given as arguments."),
    F("Math", "QUOTIENT", "numerator, denominator", "Returns the integer portion of a division."),
    F("Math", "RADIANS", "angle", "Converts degrees to radians."),
    F("Math", "RAND", "", "Returns a random number between 0 and 1."),
    F("Math", "RANDBETWEEN", "bottom, top", "Returns a random integer between the numbers you give."),
    F("Math", "ROUND", "number, num_digits", "Rounds a number to a given number of digits."),
    F("Math", "ROUNDDOWN", "number, num_digits", "Rounds a number down, toward zero."),
    F("Math", "ROUNDUP", "number, num_digits", "Rounds a number up, away from zero."),
    F("Math", "SIGN", "number", "Returns the sign of a number: 1, 0 or -1."),
    F("Math", "SIN", "number", "Returns the sine of an angle."),
    F("Math", "SQRT", "number", "Returns a positive square root."),
    F("Math", "SUBTOTAL", "function_num, ref1, [ref2], ...", "Returns a subtotal: 1 AVERAGE, 2 COUNT, 3 COUNTA, 4 MAX, 5 MIN, 6 PRODUCT, 7 STDEV, 8 STDEVP, 9 SUM, 10 VAR, 11 VARP."),
    F("Math", "SUM", "number1, [number2], ...", "Adds all the numbers in a range of cells."),
    F("Math", "SUMIF", "range, criteria, [sum_range]", "Adds the cells that meet a condition."),
    F("Math", "SUMIFS", "sum_range, criteria_range1, criteria1, ...", "Adds the cells that meet several conditions."),
    F("Math", "SUMPRODUCT", "array1, [array2], ...", "Returns the sum of the products of matching items."),
    F("Math", "SUMSQ", "number1, [number2], ...", "Returns the sum of the squares of the arguments."),
    F("Math", "TAN", "number", "Returns the tangent of an angle."),
    F("Math", "TRUNC", "number, [num_digits]", "Truncates a number to an integer or to the digits you give."),
    // Statistical
    F("Statistical", "AVERAGE", "number1, [number2], ...", "Returns the average (arithmetic mean) of its arguments."),
    F("Statistical", "AVERAGEA", "value1, [value2], ...", "Returns the average, counting text as 0 and TRUE as 1."),
    F("Statistical", "AVERAGEIF", "range, criteria, [average_range]", "Returns the average of the cells that meet a condition."),
    F("Statistical", "AVERAGEIFS", "average_range, criteria_range1, criteria1, ...", "Returns the average of the cells that meet several conditions."),
    F("Statistical", "COUNT", "value1, [value2], ...", "Counts the cells that contain numbers."),
    F("Statistical", "COUNTA", "value1, [value2], ...", "Counts the cells that are not empty."),
    F("Statistical", "COUNTBLANK", "range", "Counts the empty cells in a range."),
    F("Statistical", "COUNTIF", "range, criteria", "Counts the cells that meet a condition."),
    F("Statistical", "COUNTIFS", "criteria_range1, criteria1, ...", "Counts the cells that meet several conditions."),
    F("Statistical", "COUNTUNIQUE", "value1, [value2], ...", "Counts the different values that are not empty."),
    F("Statistical", "LARGE", "array, k", "Returns the k-th largest value."),
    F("Statistical", "MAX", "number1, [number2], ...", "Returns the largest value."),
    F("Statistical", "MAXA", "value1, [value2], ...", "Returns the largest value, counting text as 0."),
    F("Statistical", "MAXIFS", "max_range, criteria_range1, criteria1, ...", "Returns the largest value among cells that meet conditions."),
    F("Statistical", "MEDIAN", "number1, [number2], ...", "Returns the median of the numbers given."),
    F("Statistical", "MIN", "number1, [number2], ...", "Returns the smallest value."),
    F("Statistical", "MINA", "value1, [value2], ...", "Returns the smallest value, counting text as 0."),
    F("Statistical", "MINIFS", "min_range, criteria_range1, criteria1, ...", "Returns the smallest value among cells that meet conditions."),
    F("Statistical", "MODE", "number1, [number2], ...", "Returns the most common value."),
    F("Statistical", "MODE.SNGL", "number1, [number2], ...", "Returns the most common value."),
    F("Statistical", "PERCENTILE", "array, k", "Returns the k-th percentile of values."),
    F("Statistical", "PERCENTILE.INC", "array, k", "Returns the k-th percentile of values, k from 0 to 1."),
    F("Statistical", "QUARTILE", "array, quart", "Returns the quartile of a data set."),
    F("Statistical", "QUARTILE.INC", "array, quart", "Returns the quartile of a data set."),
    F("Statistical", "RANK", "number, ref, [order]", "Returns the rank of a number in a list."),
    F("Statistical", "RANK.EQ", "number, ref, [order]", "Returns the rank of a number in a list."),
    F("Statistical", "SMALL", "array, k", "Returns the k-th smallest value."),
    F("Statistical", "STDEV", "number1, [number2], ...", "Estimates standard deviation from a sample."),
    F("Statistical", "STDEV.P", "number1, [number2], ...", "Standard deviation of a whole population."),
    F("Statistical", "STDEV.S", "number1, [number2], ...", "Estimates standard deviation from a sample."),
    F("Statistical", "STDEVP", "number1, [number2], ...", "Standard deviation of a whole population."),
    F("Statistical", "VAR", "number1, [number2], ...", "Estimates variance from a sample."),
    F("Statistical", "VAR.P", "number1, [number2], ...", "Variance of a whole population."),
    F("Statistical", "VAR.S", "number1, [number2], ...", "Estimates variance from a sample."),
    F("Statistical", "VARP", "number1, [number2], ...", "Variance of a whole population."),
    // Text
    F("Text", "CHAR", "number", "Returns the character for a code number."),
    F("Text", "CLEAN", "text", "Removes all nonprintable characters from text."),
    F("Text", "CODE", "text", "Returns the code number of the first character."),
    F("Text", "CONCAT", "text1, [text2], ...", "Joins several texts into one."),
    F("Text", "CONCATENATE", "text1, [text2], ...", "Joins several texts into one."),
    F("Text", "EXACT", "text1, text2", "Checks whether two texts are exactly the same, case included."),
    F("Text", "FIND", "find_text, within_text, [start_num]", "Finds one text inside another (case-sensitive)."),
    F("Text", "FIXED", "number, [decimals], [no_commas]", "Formats a number as text with a fixed number of decimals."),
    F("Text", "LEFT", "text, [num_chars]", "Returns the leftmost characters of a text."),
    F("Text", "LEN", "text", "Returns the number of characters in a text."),
    F("Text", "LOWER", "text", "Converts text to lowercase."),
    F("Text", "MID", "text, start_num, num_chars", "Returns characters from the middle of a text."),
    F("Text", "NUMBERVALUE", "text, [decimal_separator], [group_separator]", "Converts text to a number."),
    F("Text", "PROPER", "text", "Capitalises the first letter of each word."),
    F("Text", "REPLACE", "old_text, start_num, num_chars, new_text", "Replaces characters within text."),
    F("Text", "REPT", "text, number_times", "Repeats text a given number of times."),
    F("Text", "RIGHT", "text, [num_chars]", "Returns the rightmost characters of a text."),
    F("Text", "SEARCH", "find_text, within_text, [start_num]", "Finds one text inside another (not case-sensitive, wildcards allowed)."),
    F("Text", "SUBSTITUTE", "text, old_text, new_text, [instance_num]", "Substitutes new text for old text."),
    F("Text", "T", "value", "Returns the text of a value, or empty text."),
    F("Text", "TEXT", "value, format_text", "Formats a number and converts it to text."),
    F("Text", "TEXTAFTER", "text, delimiter, [instance_num]", "Returns the text after a delimiter."),
    F("Text", "TEXTBEFORE", "text, delimiter, [instance_num]", "Returns the text before a delimiter."),
    F("Text", "TEXTJOIN", "delimiter, ignore_empty, text1, ...", "Joins texts with a delimiter."),
    F("Text", "TRIM", "text", "Removes extra spaces from text."),
    F("Text", "UNICHAR", "number", "Returns the Unicode character for a number."),
    F("Text", "UNICODE", "text", "Returns the Unicode number of the first character."),
    F("Text", "UPPER", "text", "Converts text to uppercase."),
    F("Text", "VALUE", "text", "Converts text to a number."),
    // Logical
    F("Logical", "AND", "logical1, [logical2], ...", "TRUE when all arguments are TRUE."),
    F("Logical", "FALSE", "", "Returns the logical value FALSE."),
    F("Logical", "IF", "logical_test, [value_if_true], [value_if_false]", "Returns one value if a condition is TRUE and another if FALSE."),
    F("Logical", "IFERROR", "value, value_if_error", "Returns value_if_error when the value is an error."),
    F("Logical", "IFNA", "value, value_if_na", "Returns value_if_na when the value is #N/A."),
    F("Logical", "IFS", "logical_test1, value_if_true1, ...", "Returns the value for the first TRUE condition."),
    F("Logical", "NOT", "logical", "Reverses a logical value."),
    F("Logical", "OR", "logical1, [logical2], ...", "TRUE when any argument is TRUE."),
    F("Logical", "SWITCH", "expression, value1, result1, ..., [default]", "Returns the result for the first matching value."),
    F("Logical", "TRUE", "", "Returns the logical value TRUE."),
    F("Logical", "XOR", "logical1, [logical2], ...", "TRUE when an odd number of arguments are TRUE."),
    // Lookup & reference
    F("Lookup", "CHOOSE", "index_num, value1, [value2], ...", "Chooses a value from a list by position."),
    F("Lookup", "COLUMN", "[reference]", "Returns the column number of a reference."),
    F("Lookup", "COLUMNS", "array", "Returns the number of columns in a reference."),
    F("Lookup", "HLOOKUP", "lookup_value, table_array, row_index_num, [range_lookup]", "Looks in the top row of a table and returns a value from a row you name."),
    F("Lookup", "INDEX", "array, row_num, [column_num]", "Returns a value from a range by row and column."),
    F("Lookup", "MATCH", "lookup_value, lookup_array, [match_type]", "Returns the position of a value in a range."),
    F("Lookup", "ROW", "[reference]", "Returns the row number of a reference."),
    F("Lookup", "ROWS", "array", "Returns the number of rows in a reference."),
    F("Lookup", "VLOOKUP", "lookup_value, table_array, col_index_num, [range_lookup]", "Looks in the first column of a table and returns a value from a column you name."),
    F("Lookup", "XLOOKUP", "lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]", "Searches a range and returns the matching item from another range."),
    // Date & time
    F("Date & Time", "DATE", "year, month, day", "Returns the serial number of a date."),
    F("Date & Time", "DATEDIF", "start_date, end_date, unit", "Days, months or years between two dates (unit \"D\", \"M\", \"Y\", \"YM\", \"MD\", \"YD\")."),
    F("Date & Time", "DATEVALUE", "date_text", "Converts a date in text to a serial number."),
    F("Date & Time", "DAY", "serial_number", "Returns the day of the month."),
    F("Date & Time", "DAYS", "end_date, start_date", "Returns the number of days between two dates."),
    F("Date & Time", "EDATE", "start_date, months", "Returns the date a number of months before or after."),
    F("Date & Time", "EOMONTH", "start_date, months", "Returns the last day of the month, months before or after."),
    F("Date & Time", "HOUR", "serial_number", "Returns the hour of a time value."),
    F("Date & Time", "MINUTE", "serial_number", "Returns the minute of a time value."),
    F("Date & Time", "MONTH", "serial_number", "Returns the month of a date."),
    F("Date & Time", "NETWORKDAYS", "start_date, end_date, [holidays]", "Returns the number of whole working days between two dates."),
    F("Date & Time", "NOW", "", "Returns the current date and time."),
    F("Date & Time", "SECOND", "serial_number", "Returns the second of a time value."),
    F("Date & Time", "TIME", "hour, minute, second", "Returns the serial number of a time."),
    F("Date & Time", "TODAY", "", "Returns today's date."),
    F("Date & Time", "WEEKDAY", "serial_number, [return_type]", "Returns the day of the week as a number."),
    F("Date & Time", "WEEKNUM", "serial_number, [return_type]", "Returns the week number in the year."),
    F("Date & Time", "WORKDAY", "start_date, days, [holidays]", "Returns the date a number of working days before or after."),
    F("Date & Time", "YEAR", "serial_number", "Returns the year of a date."),
    F("Date & Time", "YEARFRAC", "start_date, end_date, [basis]", "Returns the fraction of the year between two dates."),
    // Information
    F("Information", "ISBLANK", "value", "TRUE when the cell is empty."),
    F("Information", "ISERR", "value", "TRUE when the value is any error except #N/A."),
    F("Information", "ISERROR", "value", "TRUE when the value is any error."),
    F("Information", "ISEVEN", "number", "TRUE when the number is even."),
    F("Information", "ISLOGICAL", "value", "TRUE when the value is TRUE or FALSE."),
    F("Information", "ISNA", "value", "TRUE when the value is #N/A."),
    F("Information", "ISNONTEXT", "value", "TRUE when the value is not text."),
    F("Information", "ISNUMBER", "value", "TRUE when the value is a number."),
    F("Information", "ISODD", "number", "TRUE when the number is odd."),
    F("Information", "ISTEXT", "value", "TRUE when the value is text."),
    F("Information", "N", "value", "Converts a value to a number."),
    F("Information", "NA", "", "Returns the error value #N/A."),
    // Financial
    F("Financial", "FV", "rate, nper, pmt, [pv], [type]", "Returns the future value of an investment."),
    F("Financial", "IPMT", "rate, per, nper, pv, [fv], [type]", "Returns the interest payment for a period."),
    F("Financial", "NPER", "rate, pmt, pv, [fv], [type]", "Returns the number of periods for an investment."),
    F("Financial", "NPV", "rate, value1, [value2], ...", "Returns the net present value of cash flows."),
    F("Financial", "PMT", "rate, nper, pv, [fv], [type]", "Returns the payment for a loan."),
    F("Financial", "PPMT", "rate, per, nper, pv, [fv], [type]", "Returns the principal payment for a period."),
    F("Financial", "PV", "rate, nper, pmt, [fv], [type]", "Returns the present value of an investment."),
];

const BY_NAME = new Map(FUNCTIONS.map((f) => [f.name, f]));

export function fnInfo(name: string): FnInfo | undefined {
    return BY_NAME.get(name.toUpperCase());
}

/** The function name being typed at the end of a formula ("=SUM(A1,AV" gives "AV"), or null. */
export function typedFunctionPrefix(text: string, caret = text.length): string | null {
    if (!text.startsWith("=")) return null;
    const before = text.slice(0, caret);
    if (inString(before)) return null;
    const m = /(?:^=|[(,+\-*/^&=<>:; ])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
    if (!m) return null;
    // A cell reference such as A1 or AB12 is not a function name.
    if (/^[A-Za-z]{1,3}\d+$/.test(m[1])) return null;
    return m[1];
}

/** Functions whose name starts with the prefix, best first. */
export function suggest(prefix: string, limit = 12): FnInfo[] {
    const p = prefix.toUpperCase();
    return FUNCTIONS.filter((f) => f.name.startsWith(p)).slice(0, limit);
}

/** The innermost function the caret is inside, and which argument (0-based) it is on. */
export function activeCall(text: string, caret = text.length): { name: string; arg: number } | null {
    if (!text.startsWith("=")) return null;
    const stack: { name: string; arg: number }[] = [];
    let quoted = false;
    for (let i = 1; i < caret; i++) {
        const ch = text[i];
        if (ch === '"') { quoted = !quoted; continue; }
        if (quoted) continue;
        if (ch === "(") {
            const m = /([A-Za-z][A-Za-z0-9.]*)\s*$/.exec(text.slice(0, i));
            stack.push({ name: m ? m[1].toUpperCase() : "", arg: 0 });
        } else if (ch === ")") stack.pop();
        else if (ch === "," && stack.length > 0) stack[stack.length - 1].arg++;
    }
    for (let i = stack.length - 1; i >= 0; i--) if (stack[i].name && fnInfo(stack[i].name)) return stack[i];
    return null;
}

/** Replace the typed prefix at the caret with "NAME(". Returns the new text and caret. */
export function acceptSuggestion(text: string, caret: number, name: string): { text: string; caret: number } {
    const prefix = typedFunctionPrefix(text, caret) ?? "";
    const start = caret - prefix.length;
    const insert = name + "(";
    return { text: text.slice(0, start) + insert + text.slice(caret), caret: start + insert.length };
}

function inString(s: string): boolean {
    let q = false;
    for (const ch of s) if (ch === '"') q = !q;
    return q;
}
