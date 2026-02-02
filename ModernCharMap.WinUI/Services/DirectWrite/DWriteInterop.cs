using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace ModernCharMap.WinUI.Services.DirectWrite
{
    /// <summary>
    /// Provides managed entry points for creating DirectWrite COM objects and
    /// building OpenType table tag constants.
    /// </summary>
    /// <remarks>
    /// DirectWrite is the Windows text rendering and font enumeration API.
    /// This class wraps the native <c>DWriteCreateFactory</c> export from
    /// <c>dwrite.dll</c> and exposes it as a typed <see cref="IDWriteFactory"/>.
    /// </remarks>
    internal static class DWrite
    {
        /// <summary>
        /// P/Invoke for the native DirectWrite factory creation function.
        /// </summary>
        /// <param name="factoryType">
        /// The factory sharing mode. Use <c>0</c> for <c>DWRITE_FACTORY_TYPE_SHARED</c>
        /// (process-wide singleton) or <c>1</c> for <c>DWRITE_FACTORY_TYPE_ISOLATED</c>.
        /// </param>
        /// <param name="riid">The COM interface identifier for <see cref="IDWriteFactory"/>.</param>
        /// <param name="factory">Receives the factory COM object on success.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [DllImport("dwrite.dll")]
        public static extern int DWriteCreateFactory(
            int factoryType,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [MarshalAs(UnmanagedType.IUnknown)] out object factory);

        /// <summary>
        /// Creates a shared <see cref="IDWriteFactory"/> instance.
        /// </summary>
        /// <returns>A shared DirectWrite factory for font enumeration and rendering.</returns>
        /// <exception cref="COMException">Thrown if the native call fails.</exception>
        public static IDWriteFactory CreateFactory()
        {
            var iid = typeof(IDWriteFactory).GUID;
            int hr = DWriteCreateFactory(0 /* DWRITE_FACTORY_TYPE_SHARED */, iid, out object obj);
            Marshal.ThrowExceptionForHR(hr);
            return (IDWriteFactory)obj;
        }

        /// <summary>
        /// Builds an OpenType 4-byte table tag from four ASCII characters.
        /// Tags are stored in little-endian order per the OpenType specification.
        /// </summary>
        /// <param name="a">First character (lowest byte).</param>
        /// <param name="b">Second character.</param>
        /// <param name="c">Third character.</param>
        /// <param name="d">Fourth character (highest byte).</param>
        /// <returns>A 32-bit tag value suitable for <see cref="IDWriteFontFace.TryGetFontTable"/>.</returns>
        public static uint MakeTag(char a, char b, char c, char d) =>
            ((uint)d << 24) | ((uint)c << 16) | ((uint)b << 8) | (uint)a;

        /// <summary>OpenType 'post' table tag — contains PostScript glyph names.</summary>
        public static readonly uint TAG_POST = MakeTag('p', 'o', 's', 't');

        /// <summary>OpenType 'cmap' table tag — maps Unicode codepoints to glyph indices.</summary>
        public static readonly uint TAG_CMAP = MakeTag('c', 'm', 'a', 'p');
    }

    // ── IDWriteFactory ──────────────────────────────────────────────────────

    /// <summary>
    /// COM interface for the DirectWrite factory (IDWriteFactory).
    /// Provides access to the system font collection and font creation methods.
    /// </summary>
    /// <remarks>
    /// Only <see cref="GetSystemFontCollection"/> is used directly; all other
    /// vtable slots are stubbed with placeholder methods to maintain correct
    /// COM vtable layout.
    /// See: https://learn.microsoft.com/windows/win32/api/dwrite/nn-dwrite-idwritefactory
    /// </remarks>
    [ComImport, Guid("b859ee5a-d838-4b5b-a2e8-1adc7d93db48")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFactory
    {
        /// <summary>
        /// Retrieves the system font collection, optionally checking for updates.
        /// </summary>
        /// <param name="fontCollection">Receives the system font collection.</param>
        /// <param name="checkForUpdates">
        /// When <c>true</c>, DirectWrite re-enumerates installed fonts to pick up
        /// recent installs or removals.
        /// </param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int GetSystemFontCollection(out IDWriteFontCollection fontCollection,
            [MarshalAs(UnmanagedType.Bool)] bool checkForUpdates);

        // Vtable stubs for unused methods (must be present to keep COM slot ordering)
        void _s4();   // CreateCustomFontCollection
        void _s5();   // RegisterFontCollectionLoader
        void _s6();   // UnregisterFontCollectionLoader
        void _s7();   // CreateFontFileReference
        void _s8();   // CreateCustomFontFileReference
        void _s9();   // CreateFontFace
        void _s10();  // CreateRenderingParams
        void _s11();  // CreateMonitorRenderingParams
        void _s12();  // CreateCustomRenderingParams
        void _s13();  // CreateTextFormat
        void _s14();  // CreateTypography
        void _s15();  // GetGdiInterop
        void _s16();  // CreateTextLayout
        void _s17();  // CreateGdiCompatibleTextLayout
        void _s18();  // RegisterFontFileLoader
        void _s19();  // UnregisterFontFileLoader
        void _s20();  // CreateNumberSubstitution
        void _s21();  // CreateGlyphRunAnalysis
    }

    // ── IDWriteFontCollection ───────────────────────────────────────────────

    /// <summary>
    /// COM interface for a DirectWrite font collection (IDWriteFontCollection).
    /// Represents all font families available in the system or a custom collection.
    /// </summary>
    /// <remarks>
    /// See: https://learn.microsoft.com/windows/win32/api/dwrite/nn-dwrite-idwritefontcollection
    /// </remarks>
    [ComImport, Guid("a84cee02-3eea-4eee-a827-87c1a02a0fcc")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFontCollection
    {
        /// <summary>Returns the number of font families in this collection.</summary>
        [PreserveSig] uint GetFontFamilyCount();

        /// <summary>
        /// Retrieves a font family by its zero-based index.
        /// </summary>
        /// <param name="index">Zero-based index within the collection.</param>
        /// <param name="fontFamily">Receives the font family object.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int GetFontFamily(uint index, out IDWriteFontFamily fontFamily);

        /// <summary>
        /// Looks up a font family by name, returning its index if found.
        /// </summary>
        /// <param name="familyName">The family name to search for (case-insensitive).</param>
        /// <param name="index">Receives the zero-based index if found.</param>
        /// <param name="exists">Receives <c>true</c> if the family was found.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int FindFamilyName(
            [MarshalAs(UnmanagedType.LPWStr)] string familyName,
            out uint index,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        void _s6(); // GetFontFromFontFace
    }

    // ── IDWriteFontFamily (inherits IDWriteFontList) ────────────────────────

    /// <summary>
    /// COM interface for a DirectWrite font family (IDWriteFontFamily).
    /// Inherits from IDWriteFontList, adding the ability to retrieve localized
    /// family names.
    /// </summary>
    /// <remarks>
    /// See: https://learn.microsoft.com/windows/win32/api/dwrite/nn-dwrite-idwritefontfamily
    /// </remarks>
    [ComImport, Guid("da20d8ef-812a-4c43-9802-62ec4abd7add")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFontFamily
    {
        void _s3(); // GetFontCollection (from IDWriteFontList)

        /// <summary>Returns the number of fonts in this family (e.g. Regular, Bold, Italic).</summary>
        [PreserveSig] uint GetFontCount();

        /// <summary>
        /// Retrieves a specific font by index within the family.
        /// </summary>
        /// <param name="index">Zero-based font index.</param>
        /// <param name="font">Receives the font object.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int GetFont(uint index, out IDWriteFont font);

        /// <summary>
        /// Retrieves the localized family names (e.g. "Arial" in English, localized equivalents in other cultures).
        /// </summary>
        /// <param name="names">Receives the localized string set.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int GetFamilyNames(out IDWriteLocalizedStrings names);

        void _s7(); // GetFirstMatchingFont
        void _s8(); // GetMatchingFonts
    }

    // ── IDWriteFont ─────────────────────────────────────────────────────────

    /// <summary>
    /// COM interface for a single DirectWrite font (IDWriteFont).
    /// Represents one face within a family (e.g. "Arial Bold") and provides
    /// metadata such as weight, style, and whether it is a symbol font.
    /// </summary>
    /// <remarks>
    /// See: https://learn.microsoft.com/windows/win32/api/dwrite/nn-dwrite-idwritefont
    /// </remarks>
    [ComImport, Guid("acd16696-8c14-4f5d-877e-fe3fc1d32737")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFont
    {
        void _s3(); // GetFontFamily

        /// <summary>Returns the font weight (e.g. 400 for Regular, 700 for Bold).</summary>
        [PreserveSig] int GetWeight();

        /// <summary>Returns the font stretch value (e.g. condensed, normal, expanded).</summary>
        [PreserveSig] int GetStretch();

        /// <summary>Returns the font style (normal, oblique, or italic).</summary>
        [PreserveSig] int GetStyle();

        /// <summary>
        /// Indicates whether this font uses a symbol encoding rather than Unicode.
        /// </summary>
        /// <returns><c>true</c> if the font is a symbol font; otherwise <c>false</c>.</returns>
        [PreserveSig]
        [return: MarshalAs(UnmanagedType.Bool)]
        bool IsSymbolFont();

        /// <summary>Retrieves the localized face names (e.g. "Regular", "Bold Italic").</summary>
        [PreserveSig]
        int GetFaceNames(out IDWriteLocalizedStrings names);

        /// <summary>
        /// Retrieves an informational string such as the font designer or license URL.
        /// </summary>
        /// <param name="informationalStringID">The type of informational string to retrieve.</param>
        /// <param name="informationalStrings">Receives the localized strings if found.</param>
        /// <param name="exists">Receives <c>true</c> if the requested string type exists.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int GetInformationalStrings(
            int informationalStringID,
            out IDWriteLocalizedStrings informationalStrings,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        /// <summary>Returns the font simulation flags (none, bold, oblique).</summary>
        [PreserveSig] int GetSimulations();
        void _s11(); // GetMetrics

        /// <summary>
        /// Checks whether the font supports a specific Unicode codepoint.
        /// </summary>
        /// <param name="unicodeValue">The Unicode codepoint to test.</param>
        /// <param name="exists">Receives <c>true</c> if the font contains a glyph for this codepoint.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int HasCharacter(uint unicodeValue,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        /// <summary>
        /// Creates an <see cref="IDWriteFontFace"/> for accessing glyph-level data
        /// such as glyph indices, font tables, and outline geometry.
        /// </summary>
        /// <param name="fontFace">Receives the font face object.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int CreateFontFace(out IDWriteFontFace fontFace);
    }

    // ── IDWriteFontFace ─────────────────────────────────────────────────────

    /// <summary>
    /// COM interface for a DirectWrite font face (IDWriteFontFace).
    /// Provides glyph-level access: mapping codepoints to glyph indices,
    /// reading OpenType tables, and retrieving glyph metrics.
    /// </summary>
    /// <remarks>
    /// This interface can be cast (via COM QueryInterface) to
    /// <see cref="IDWriteFontFace1"/> for additional capabilities such as
    /// full-Unicode range enumeration.
    /// See: https://learn.microsoft.com/windows/win32/api/dwrite/nn-dwrite-idwritefontface
    /// </remarks>
    [ComImport, Guid("5f49804d-7024-4d43-bfa9-d25984f53849")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFontFace
    {
        void _s3();  // GetType
        void _s4();  // GetFiles
        void _s5();  // GetIndex
        void _s6();  // GetSimulations
        void _s7();  // IsSymbolFont
        void _s8();  // GetMetrics
        void _s9();  // GetGlyphCount
        void _s10(); // GetDesignGlyphMetrics

        /// <summary>
        /// Maps an array of Unicode codepoints to their corresponding glyph indices.
        /// A glyph index of <c>0</c> means the font has no glyph for that codepoint.
        /// </summary>
        /// <param name="codePoints">Array of Unicode codepoints (as <c>uint</c>).</param>
        /// <param name="codePointCount">Number of codepoints in the array.</param>
        /// <param name="glyphIndices">Receives the corresponding glyph indices.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int GetGlyphIndices(
            [In, MarshalAs(UnmanagedType.LPArray)] uint[] codePoints,
            uint codePointCount,
            [Out, MarshalAs(UnmanagedType.LPArray)] ushort[] glyphIndices);

        /// <summary>
        /// Reads a raw OpenType table from the font file.
        /// The returned pointer is valid until <see cref="ReleaseFontTable"/> is called.
        /// </summary>
        /// <param name="openTypeTableTag">
        /// The 4-byte table tag (use <see cref="DWrite.MakeTag"/> to build).
        /// </param>
        /// <param name="tableData">Receives a pointer to the table data (read-only).</param>
        /// <param name="tableSize">Receives the table size in bytes.</param>
        /// <param name="tableContext">Receives an opaque context for <see cref="ReleaseFontTable"/>.</param>
        /// <param name="exists">Receives <c>true</c> if the table was found.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int TryGetFontTable(
            uint openTypeTableTag,
            out IntPtr tableData,
            out uint tableSize,
            out IntPtr tableContext,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        /// <summary>
        /// Releases a font table pointer previously obtained via <see cref="TryGetFontTable"/>.
        /// </summary>
        /// <param name="tableContext">The opaque context returned by TryGetFontTable.</param>
        void ReleaseFontTable(IntPtr tableContext);

        void _s14(); // GetGlyphRunOutline
        void _s15(); // GetRecommendedRenderingMode
        void _s16(); // GetGdiCompatibleMetrics
        void _s17(); // GetGdiCompatibleGlyphMetrics
    }

    // ── IDWriteFontFace1 ────────────────────────────────────────────────────

    /// <summary>
    /// COM interface extending <see cref="IDWriteFontFace"/> with full-Unicode
    /// range enumeration via <see cref="GetUnicodeRanges"/> (IDWriteFontFace1).
    /// Available on Windows 8 and later.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Because COM interfaces use vtable-based dispatch, this managed interface
    /// must declare stub methods for every inherited slot from
    /// <see cref="IDWriteFontFace"/> (15 slots) plus the three new
    /// IDWriteFontFace1 methods that precede <see cref="GetUnicodeRanges"/>
    /// in the vtable. Omitting any slot would cause all subsequent method
    /// calls to invoke the wrong native function.
    /// </para>
    /// <para>
    /// Unlike the GDI <c>GetFontUnicodeRanges</c> function (which uses 16-bit
    /// <c>WCRANGE</c> and is limited to the Basic Multilingual Plane),
    /// <see cref="GetUnicodeRanges"/> uses 32-bit <see cref="DWriteUnicodeRange"/>
    /// structs and supports the full Unicode range including Supplementary
    /// Multilingual Plane characters (emoji, symbols above U+FFFF).
    /// </para>
    /// <para>
    /// See: https://learn.microsoft.com/windows/win32/api/dwrite_1/nn-dwrite_1-idwritefontface1
    /// </para>
    /// </remarks>
    [ComImport, Guid("a71efdb4-9fdb-4838-ad90-cfc3be8c3daf")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFontFace1
    {
        // --- IDWriteFontFace inherited methods (15 vtable slots) ---
        void _ff_GetType();
        void _ff_GetFiles();
        void _ff_GetIndex();
        void _ff_GetSimulations();
        void _ff_IsSymbolFont();
        void _ff_GetMetrics();
        void _ff_GetGlyphCount();
        void _ff_GetDesignGlyphMetrics();
        void _ff_GetGlyphIndices();
        void _ff_TryGetFontTable();
        void _ff_ReleaseFontTable();
        void _ff_GetGlyphRunOutline();
        void _ff_GetRecommendedRenderingMode();
        void _ff_GetGdiCompatibleMetrics();
        void _ff_GetGdiCompatibleGlyphMetrics();

        // --- IDWriteFontFace1 new methods (3 vtable slots before GetUnicodeRanges) ---
        void _ff1_GetMetrics1();
        void _ff1_GetGdiCompatibleMetrics1();
        void _ff1_GetCaretMetrics();

        /// <summary>
        /// Retrieves the Unicode codepoint ranges supported by this font face.
        /// Uses a two-call pattern: call once with <paramref name="maxRangeCount"/> = 0
        /// to get the required count, then allocate and call again.
        /// </summary>
        /// <param name="maxRangeCount">
        /// Maximum number of ranges the buffer can hold, or <c>0</c> to query the count.
        /// </param>
        /// <param name="unicodeRanges">
        /// Pointer to a buffer of <see cref="DWriteUnicodeRange"/> structs, or
        /// <see cref="IntPtr.Zero"/> when querying the count.
        /// </param>
        /// <param name="actualRangeCount">Receives the actual number of ranges.</param>
        /// <returns>
        /// <c>S_OK</c> on success, or <c>E_NOT_SUFFICIENT_BUFFER</c> when the
        /// buffer is too small (which is expected on the first sizing call).
        /// </returns>
        [PreserveSig]
        int GetUnicodeRanges(
            uint maxRangeCount,
            IntPtr unicodeRanges,
            out uint actualRangeCount);
    }

    /// <summary>
    /// Represents a contiguous range of Unicode codepoints supported by a font.
    /// Mirrors the native <c>DWRITE_UNICODE_RANGE</c> structure.
    /// </summary>
    /// <remarks>
    /// Both <see cref="First"/> and <see cref="Last"/> are inclusive, and use
    /// 32-bit unsigned integers to support the full Unicode range (U+0000 through U+10FFFF).
    /// </remarks>
    [StructLayout(LayoutKind.Sequential)]
    internal struct DWriteUnicodeRange
    {
        /// <summary>The first Unicode codepoint in the range (inclusive).</summary>
        public uint First;

        /// <summary>The last Unicode codepoint in the range (inclusive).</summary>
        public uint Last;
    }

    // ── IDWriteLocalizedStrings ─────────────────────────────────────────────

    /// <summary>
    /// COM interface for a set of locale-specific strings (IDWriteLocalizedStrings).
    /// Used to retrieve font family names, face names, and informational strings
    /// in the user's preferred language.
    /// </summary>
    /// <remarks>
    /// See: https://learn.microsoft.com/windows/win32/api/dwrite/nn-dwrite-idwritelocalizedstrings
    /// </remarks>
    [ComImport, Guid("08256209-099a-4b34-b86d-c22b110e7771")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteLocalizedStrings
    {
        /// <summary>Returns the number of locale/string pairs.</summary>
        [PreserveSig] uint GetCount();

        /// <summary>
        /// Finds the index of a string for the specified locale name (e.g. "en-us").
        /// </summary>
        /// <param name="localeName">The BCP 47 locale name to search for.</param>
        /// <param name="index">Receives the zero-based index if found.</param>
        /// <param name="exists">Receives <c>true</c> if the locale was found.</param>
        /// <returns>An HRESULT indicating success or failure.</returns>
        [PreserveSig]
        int FindLocaleName(
            [MarshalAs(UnmanagedType.LPWStr)] string localeName,
            out uint index,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        /// <summary>Gets the length (in characters, excluding null terminator) of a locale name at the given index.</summary>
        [PreserveSig]
        int GetLocaleNameLength(uint index, out uint length);

        /// <summary>Copies the locale name at the given index into the provided buffer.</summary>
        [PreserveSig]
        int GetLocaleName(uint index,
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder localeName,
            uint size);

        /// <summary>Gets the length (in characters, excluding null terminator) of the string at the given index.</summary>
        [PreserveSig]
        int GetStringLength(uint index, out uint length);

        /// <summary>Copies the string at the given index into the provided buffer.</summary>
        [PreserveSig]
        int GetString(uint index,
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder stringBuffer,
            uint size);
    }

    // ── Helpers ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Utility methods for working with DirectWrite COM objects and reading
    /// big-endian binary data from OpenType font tables.
    /// </summary>
    internal static class DWriteHelper
    {
        /// <summary>
        /// Retrieves a string from an <see cref="IDWriteLocalizedStrings"/> set by index.
        /// </summary>
        /// <param name="strings">The localized string set.</param>
        /// <param name="index">Zero-based index of the string to retrieve.</param>
        /// <returns>The string value at the given index.</returns>
        public static string GetString(IDWriteLocalizedStrings strings, uint index)
        {
            strings.GetStringLength(index, out uint length);
            var sb = new StringBuilder((int)length + 1);
            strings.GetString(index, sb, length + 1);
            return sb.ToString();
        }

        /// <summary>
        /// Retrieves the best-matching localized string for the current UI culture.
        /// Falls back to "en-us", then to the first available string.
        /// </summary>
        /// <param name="strings">The localized string set to search.</param>
        /// <returns>The best available localized string.</returns>
        public static string GetBestLocalizedString(IDWriteLocalizedStrings strings)
        {
            strings.FindLocaleName(CultureInfo.CurrentCulture.Name, out uint index, out bool exists);
            if (!exists)
                strings.FindLocaleName("en-us", out index, out exists);
            if (!exists)
                index = 0;
            return GetString(strings, index);
        }

        /// <summary>
        /// Reads a 16-bit unsigned integer in big-endian byte order from unmanaged memory.
        /// OpenType font tables use big-endian (network) byte order.
        /// </summary>
        /// <param name="ptr">Base pointer to the memory block.</param>
        /// <param name="offset">Byte offset from <paramref name="ptr"/>.</param>
        /// <returns>The 16-bit value in host byte order.</returns>
        public static ushort ReadUInt16BE(IntPtr ptr, int offset)
        {
            byte b1 = Marshal.ReadByte(ptr, offset);
            byte b2 = Marshal.ReadByte(ptr, offset + 1);
            return (ushort)((b1 << 8) | b2);
        }

        /// <summary>
        /// Reads a 32-bit unsigned integer in big-endian byte order from unmanaged memory.
        /// OpenType font tables use big-endian (network) byte order.
        /// </summary>
        /// <param name="ptr">Base pointer to the memory block.</param>
        /// <param name="offset">Byte offset from <paramref name="ptr"/>.</param>
        /// <returns>The 32-bit value in host byte order.</returns>
        public static uint ReadUInt32BE(IntPtr ptr, int offset)
        {
            byte b1 = Marshal.ReadByte(ptr, offset);
            byte b2 = Marshal.ReadByte(ptr, offset + 1);
            byte b3 = Marshal.ReadByte(ptr, offset + 2);
            byte b4 = Marshal.ReadByte(ptr, offset + 3);
            return (uint)((b1 << 24) | (b2 << 16) | (b3 << 8) | b4);
        }
    }
}
