using System.Globalization;
using System.Text;

namespace FastDocView.Odt;

/// <summary>
/// The values of a document's sequence fields, variables and user fields at one point of the body (one name space, as in
/// LibreOffice's SwCalc), plus the sequence declarations. Index time records it at each block start (in the scan state, so a
/// span rescan gives every block's); the renderer replays the block's own fields from it in document order, so a sequence or a
/// variable shows what LibreOffice computes when it prints (SwDocUpdateField / DocumentFieldsManager::UpdateExpFieldsImpl:
/// every SetExp field in body order, a sequence restarting at a new outline node of its level). Immutable: a change makes a copy.
/// </summary>
internal sealed class OdtFieldState
{
    public static readonly OdtFieldState Empty = new(new Dictionary<string, OdtFieldValue>(StringComparer.OrdinalIgnoreCase),
                                                      new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase));

    private readonly Dictionary<string, OdtFieldValue> _values;
    /// <summary>Sequence name -> text:display-outline-level (1-based, 0 = none; ODF 1.3 part 3 7.4.13.1).</summary>
    private readonly Dictionary<string, int> _sequenceLevels;

    private OdtFieldState(Dictionary<string, OdtFieldValue> values, Dictionary<string, int> sequenceLevels)
    {
        _values = values;
        _sequenceLevels = sequenceLevels;
    }

    public bool IsEmpty => _values.Count == 0 && _sequenceLevels.Count == 0;

    public OdtFieldValue? ValueOf(string name) => _values.TryGetValue(name, out OdtFieldValue value) ? value : null;

    private OdtFieldState Copy() => new(new Dictionary<string, OdtFieldValue>(_values, StringComparer.OrdinalIgnoreCase),
                                        new Dictionary<string, int>(_sequenceLevels, StringComparer.OrdinalIgnoreCase));

    /// <summary>text:sequence-decl (7.4.13.1): the counter starts at 0; display-outline-level restarts it per chapter.</summary>
    public OdtFieldState Declare(string name, string? outlineLevel)
    {
        int level = int.TryParse(outlineLevel, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed is > 0 and <= 10 ? parsed : 0;
        if (_sequenceLevels.TryGetValue(name, out int known) && known == level) return this;
        OdtFieldState copy = Copy();
        copy._sequenceLevels[name] = level;
        return copy;
    }

    /// <summary>text:user-field-decl (7.4.7): a document-wide value - office:formula / office:value / office:string-value.</summary>
    public OdtFieldState DeclareUserField(string name, string? valueType, string? formula, string? value, string? stringValue, string? booleanValue)
    {
        OdtFieldValue? result = valueType == "string" ? OdtFieldValue.OfText(stringValue ?? "") : NumberOf(formula, value, booleanValue, null);
        if (result is null) return this;
        OdtFieldState copy = Copy();
        copy._values[name] = result.Value;
        return copy;
    }

    /// <summary>
    /// A heading of outline <paramref name="level"/> starts a new chapter for every sequence numbered within a chapter of that level
    /// or deeper (UpdateExpFieldsImpl: FindOutlineNodeOfLevel changes -> the counter is set to 0 before the next sequence).
    /// </summary>
    public OdtFieldState HeadingStarted(int level)
    {
        OdtFieldState? copy = null;
        foreach (KeyValuePair<string, int> sequence in _sequenceLevels)
            if (sequence.Value > 0 && level <= sequence.Value && (!_values.TryGetValue(sequence.Key, out OdtFieldValue value) || value.Number != 0 || value.IsText))
                (copy ??= Copy())._values[sequence.Key] = OdtFieldValue.Of(0);
        return copy ?? this;
    }

    /// <summary>text:sequence (7.4.13): the counter becomes text:formula (default "name+1"); null when the formula does not compute.</summary>
    public OdtFieldState Sequence(string name, string? formula, out OdtFieldValue? value)
    {
        value = OdtFieldFormula.Evaluate(string.IsNullOrEmpty(formula) ? name + "+1" : formula!, ValueOf);
        if (value is null) return this;
        value = OdtFieldValue.Of(value.Value.AsNumber);
        OdtFieldState copy = Copy();
        copy._values[name] = value.Value;
        return copy;
    }

    /// <summary>
    /// text:variable-set (7.4.4): a string variable takes its office:string-value or its text; any other type takes text:formula when
    /// it computes, else office:value (LibreOffice's SwSetExpField keeps the formula and recomputes it).
    /// </summary>
    public OdtFieldState SetVariable(string name, string? valueType, string? formula, string? value, string? stringValue, string? booleanValue, string? content, out OdtFieldValue? result)
    {
        result = valueType == "string" ? OdtFieldValue.OfText(stringValue ?? content ?? "") : NumberOf(formula, value, booleanValue, content);
        if (result is null) return this;
        OdtFieldState copy = Copy();
        copy._values[name] = result.Value;
        return copy;
    }

    private OdtFieldValue? NumberOf(string? formula, string? value, string? booleanValue, string? content)
    {
        if (!string.IsNullOrEmpty(formula) && OdtFieldFormula.Evaluate(formula!, ValueOf) is { } computed) return OdtFieldValue.Of(computed.AsNumber);
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return OdtFieldValue.Of(number);
        if (booleanValue is not null) return OdtFieldValue.Of(booleanValue == "true" ? 1 : 0);
        if (double.TryParse(content, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return OdtFieldValue.Of(number);
        return null;
    }

    // -- serialisation (part of the scan state) --------------------------------------------

    public void Write(List<byte> bytes)
    {
        IndexBinaryArrays.AppendVarint(bytes, (uint)_values.Count);
        foreach (KeyValuePair<string, OdtFieldValue> entry in _values.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            AppendString(bytes, entry.Key);
            if (entry.Value.IsText) { bytes.Add(1); AppendString(bytes, entry.Value.Text!); }
            else { bytes.Add(0); bytes.AddRange(BitConverter.GetBytes(entry.Value.Number)); }
        }
        IndexBinaryArrays.AppendVarint(bytes, (uint)_sequenceLevels.Count);
        foreach (KeyValuePair<string, int> entry in _sequenceLevels.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            AppendString(bytes, entry.Key);
            bytes.Add((byte)entry.Value);
        }
    }

    public static OdtFieldState Read(byte[] bytes)
    {
        if (bytes.Length == 0) return Empty;
        int position = 0;
        var values = new Dictionary<string, OdtFieldValue>(StringComparer.OrdinalIgnoreCase);
        int count = Count(bytes, ref position);
        for (int i = 0; i < count; i++)
        {
            string name = ReadString(bytes, ref position);
            if (position >= bytes.Length) throw Damaged();
            if (bytes[position++] == 1) values[name] = OdtFieldValue.OfText(ReadString(bytes, ref position));
            else
            {
                if (position + 8 > bytes.Length) throw Damaged();
                values[name] = OdtFieldValue.Of(BitConverter.ToDouble(bytes, position));
                position += 8;
            }
        }
        var levels = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        count = Count(bytes, ref position);
        for (int i = 0; i < count; i++)
        {
            string name = ReadString(bytes, ref position);
            if (position >= bytes.Length) throw Damaged();
            levels[name] = bytes[position++];
        }
        if (position != bytes.Length) throw Damaged();
        return new OdtFieldState(values, levels);
    }

    private static int Count(byte[] bytes, ref int position)
    {
        ulong count = IndexBinaryArrays.ReadVarint(bytes, ref position);
        if (count > (ulong)(bytes.Length - position)) throw Damaged();
        return (int)count;
    }

    private static void AppendString(List<byte> bytes, string text)
    {
        byte[] utf8 = Encoding.UTF8.GetBytes(text);
        IndexBinaryArrays.AppendVarint(bytes, (uint)utf8.Length);
        bytes.AddRange(utf8);
    }

    private static string ReadString(byte[] bytes, ref int position)
    {
        int length = Count(bytes, ref position);
        string text = Encoding.UTF8.GetString(bytes, position, length);
        position += length;
        return text;
    }

    private static InvalidDataException Damaged() => new("OpenDocument index: field state damaged.");
}
