using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace ModernCharMap.WinUI.Services.DirectWrite
{
    internal static class DWrite
    {
        [DllImport("dwrite.dll")]
        public static extern int DWriteCreateFactory(
            int factoryType,
            [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [MarshalAs(UnmanagedType.IUnknown)] out object factory);

        public static IDWriteFactory CreateFactory()
        {
            var iid = typeof(IDWriteFactory).GUID;
            int hr = DWriteCreateFactory(0 /* DWRITE_FACTORY_TYPE_SHARED */, iid, out object obj);
            Marshal.ThrowExceptionForHR(hr);
            return (IDWriteFactory)obj;
        }

        public static uint MakeTag(char a, char b, char c, char d) =>
            ((uint)d << 24) | ((uint)c << 16) | ((uint)b << 8) | (uint)a;

        public static readonly uint TAG_POST = MakeTag('p', 'o', 's', 't');
        public static readonly uint TAG_CMAP = MakeTag('c', 'm', 'a', 'p');
    }

    // ── IDWriteFactory ──────────────────────────────────────────────────────
    [ComImport, Guid("b859ee5a-d838-4b5b-a2e8-1adc7d93db48")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFactory
    {
        [PreserveSig]
        int GetSystemFontCollection(out IDWriteFontCollection fontCollection,
            [MarshalAs(UnmanagedType.Bool)] bool checkForUpdates);

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
    [ComImport, Guid("a84cee02-3eea-4eee-a827-87c1a02a0fcc")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFontCollection
    {
        [PreserveSig] uint GetFontFamilyCount();

        [PreserveSig]
        int GetFontFamily(uint index, out IDWriteFontFamily fontFamily);

        [PreserveSig]
        int FindFamilyName(
            [MarshalAs(UnmanagedType.LPWStr)] string familyName,
            out uint index,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        void _s6(); // GetFontFromFontFace
    }

    // ── IDWriteFontFamily (inherits IDWriteFontList) ────────────────────────
    [ComImport, Guid("da20d8ef-812a-4c43-9802-62ec4abd7add")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFontFamily
    {
        void _s3(); // GetFontCollection (from IDWriteFontList)

        [PreserveSig] uint GetFontCount();

        [PreserveSig]
        int GetFont(uint index, out IDWriteFont font);

        [PreserveSig]
        int GetFamilyNames(out IDWriteLocalizedStrings names);

        void _s7(); // GetFirstMatchingFont
        void _s8(); // GetMatchingFonts
    }

    // ── IDWriteFont ─────────────────────────────────────────────────────────
    [ComImport, Guid("acd16696-8c14-4f5d-877e-fe3fc1d32737")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteFont
    {
        void _s3(); // GetFontFamily

        [PreserveSig] int GetWeight();
        [PreserveSig] int GetStretch();
        [PreserveSig] int GetStyle();

        [PreserveSig]
        [return: MarshalAs(UnmanagedType.Bool)]
        bool IsSymbolFont();

        [PreserveSig]
        int GetFaceNames(out IDWriteLocalizedStrings names);

        [PreserveSig]
        int GetInformationalStrings(
            int informationalStringID,
            out IDWriteLocalizedStrings informationalStrings,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        [PreserveSig] int GetSimulations();
        void _s11(); // GetMetrics

        [PreserveSig]
        int HasCharacter(uint unicodeValue,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        [PreserveSig]
        int CreateFontFace(out IDWriteFontFace fontFace);
    }

    // ── IDWriteFontFace ─────────────────────────────────────────────────────
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

        [PreserveSig]
        int GetGlyphIndices(
            [In, MarshalAs(UnmanagedType.LPArray)] uint[] codePoints,
            uint codePointCount,
            [Out, MarshalAs(UnmanagedType.LPArray)] ushort[] glyphIndices);

        [PreserveSig]
        int TryGetFontTable(
            uint openTypeTableTag,
            out IntPtr tableData,
            out uint tableSize,
            out IntPtr tableContext,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        void ReleaseFontTable(IntPtr tableContext);

        void _s14(); // GetGlyphRunOutline
        void _s15(); // GetRecommendedRenderingMode
        void _s16(); // GetGdiCompatibleMetrics
        void _s17(); // GetGdiCompatibleGlyphMetrics
    }

    // ── IDWriteLocalizedStrings ─────────────────────────────────────────────
    [ComImport, Guid("08256209-099a-4b34-b86d-c22b110e7771")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDWriteLocalizedStrings
    {
        [PreserveSig] uint GetCount();

        [PreserveSig]
        int FindLocaleName(
            [MarshalAs(UnmanagedType.LPWStr)] string localeName,
            out uint index,
            [MarshalAs(UnmanagedType.Bool)] out bool exists);

        [PreserveSig]
        int GetLocaleNameLength(uint index, out uint length);

        [PreserveSig]
        int GetLocaleName(uint index,
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder localeName,
            uint size);

        [PreserveSig]
        int GetStringLength(uint index, out uint length);

        [PreserveSig]
        int GetString(uint index,
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder stringBuffer,
            uint size);
    }

    // ── Helpers ─────────────────────────────────────────────────────────────
    internal static class DWriteHelper
    {
        public static string GetString(IDWriteLocalizedStrings strings, uint index)
        {
            strings.GetStringLength(index, out uint length);
            var sb = new StringBuilder((int)length + 1);
            strings.GetString(index, sb, length + 1);
            return sb.ToString();
        }

        public static string GetBestLocalizedString(IDWriteLocalizedStrings strings)
        {
            strings.FindLocaleName(CultureInfo.CurrentCulture.Name, out uint index, out bool exists);
            if (!exists)
                strings.FindLocaleName("en-us", out index, out exists);
            if (!exists)
                index = 0;
            return GetString(strings, index);
        }

        public static ushort ReadUInt16BE(IntPtr ptr, int offset)
        {
            byte b1 = Marshal.ReadByte(ptr, offset);
            byte b2 = Marshal.ReadByte(ptr, offset + 1);
            return (ushort)((b1 << 8) | b2);
        }

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
