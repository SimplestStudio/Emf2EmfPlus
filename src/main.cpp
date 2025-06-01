#include <Windows.h>
#include <gdiplus.h>
#include <cwchar>

using namespace Gdiplus;

static const wchar_t pHelp[] = LR"(
Emf2EmfPlus - EMF to EMF+ Converter (v1.0)
===========================================
Converts standard Enhanced Metafile (EMF) files to EMF+ format using GDI+

SYNOPSIS:
  emf2emfplus.exe <input_file> <output_file>

ARGUMENTS:
  <input_file>    Path to the source EMF file.
  <output_file>   Destination path for the converted EMF+ file.

EXAMPLE:
  emf2emfplus.exe input.emf output.emf

NOTES:
  - Overwrites the output file if it already exists.
  - The resulting file will be in EMF+ format only (no dual-mode EMF).
  - Make sure the input file is a valid EMF file.
)";


static const wchar_t* GetGdiPlusError(Gdiplus::Status status)
{
    switch (status) {
    case Gdiplus::Ok: return L"Success";
    case Gdiplus::GenericError: return L"Generic error";
    case Gdiplus::InvalidParameter: return L"Invalid parameter";
    case Gdiplus::OutOfMemory: return L"Out of memory";
    case Gdiplus::ObjectBusy: return L"Object busy";
    case Gdiplus::InsufficientBuffer: return L"Insufficient buffer";
    case Gdiplus::NotImplemented: return L"Not implemented";
    case Gdiplus::Win32Error: return L"Win32 error";
    case Gdiplus::WrongState: return L"Wrong state";
    case Gdiplus::Aborted: return L"Aborted";
    case Gdiplus::FileNotFound: return L"File not found";
    case Gdiplus::ValueOverflow: return L"Value overflow";
    case Gdiplus::AccessDenied: return L"Access denied";
    case Gdiplus::UnknownImageFormat: return L"Unknown image format";
    case Gdiplus::FontFamilyNotFound: return L"Font family not found";
    case Gdiplus::FontStyleNotFound: return L"Font style not found";
    case Gdiplus::NotTrueTypeFont: return L"Not a TrueType font";
    case Gdiplus::UnsupportedGdiplusVersion: return L"Unsupported GDI+ version";
    case Gdiplus::GdiplusNotInitialized: return L"GDI+ not initialized";
    case Gdiplus::PropertyNotFound: return L"Property not found";
    case Gdiplus::PropertyNotSupported: return L"Property not supported";
    default: return L"Unknown error";
    }
}

static BOOL __SetProcessDPIAware()
{
    BOOL result = FALSE;
    if (HMODULE hGdiplus = LoadLibrary(L"user32.dll")) {
        BOOL (WINAPI *_SetProcessDPIAware)();
        *(FARPROC*)&_SetProcessDPIAware = GetProcAddress(hGdiplus, "SetProcessDPIAware");
        if (_SetProcessDPIAware) {
            result = _SetProcessDPIAware();
        }
        FreeLibrary(hGdiplus);
    }
    return result;
}

static GpStatus __ConvertToEmfPlusToFile(GpGraphics *gpGr, GpMetafile *gpEmf, LPCWSTR outEmfFilePath)
{
    GpStatus status = GpStatus::GdiplusNotInitialized;
    if (HMODULE hGdiplus = LoadLibrary(L"gdiplus.dll")) {
        GpStatus (WINGDIPAPI *_ConvertToEmfPlusToFile)(const GpGraphics*, GpMetafile*, BOOL*, LPCWSTR, Gdiplus::EmfType, LPCWSTR, GpMetafile**);
        *(FARPROC*)&_ConvertToEmfPlusToFile = GetProcAddress(hGdiplus, "GdipConvertToEmfPlusToFile");
        if (_ConvertToEmfPlusToFile) {
            BOOL success = FALSE;
            GpMetafile *gpEmfOut = nullptr;
            status = _ConvertToEmfPlusToFile(gpGr, gpEmf, &success, outEmfFilePath, EmfTypeEmfPlusOnly, L"EMFplus", &gpEmfOut);
            if (status == GpStatus::Ok /*&& success*/) {
                Gdiplus::DllExports::GdipDisposeImage((GpImage*)gpEmfOut);
            }
        }
        FreeLibrary(hGdiplus);
    }
    return status;
}

static void ConvertAndSaveEmfToEmfPlus(LPCWSTR inEmfFilePath, LPCWSTR outEmfFilePath)
{
    HDC scrDC = GetDC(nullptr);

    GpMetafile *gpEmf = nullptr;
    GpStatus status = DllExports::GdipCreateMetafileFromFile(inEmfFilePath, &gpEmf);
    if (status != GpStatus::Ok) {
        gpEmf = nullptr;
        wprintf(L"[ERROR] Failed to open input file: %s.\n", GetGdiPlusError(status));
    }

    GpGraphics *gpGr = nullptr;
    status = DllExports::GdipCreateFromHDC(scrDC, &gpGr);
    if (status == GpStatus::Ok) {
        DllExports::GdipSetPageUnit(gpGr, GpUnit::UnitPixel);
        DllExports::GdipSetSmoothingMode(gpGr, Gdiplus::SmoothingModeHighQuality);
        DllExports::GdipSetInterpolationMode(gpGr, Gdiplus::InterpolationModeHighQuality);
        DllExports::GdipSetPixelOffsetMode(gpGr, Gdiplus::PixelOffsetModeHighQuality);
    } else {
        gpGr = nullptr;
        wprintf(L"[ERROR] Failed to initialize GDI+: %s.\n", GetGdiPlusError(status));
    }

    if (gpEmf && gpGr) {
        status = __ConvertToEmfPlusToFile(gpGr, gpEmf, outEmfFilePath);
        if (status == GpStatus::Ok) {
            wprintf(L"[OK] Conversion succeeded: %s\n", outEmfFilePath);
        } else {
            wprintf(L"[ERROR] Conversion failed: %s.\n", GetGdiPlusError(status));
        }
    }

    if (gpGr)
        Gdiplus::DllExports::GdipDeleteGraphics(gpGr);
    if (gpEmf)
        Gdiplus::DllExports::GdipDisposeImage((GpImage*)gpEmf);

    ReleaseDC(nullptr, scrDC);
}

int WINAPI wmain(int argc, wchar_t *argv[])
{
    __SetProcessDPIAware();
    if (argc < 3) {
        wprintf(pHelp);
        return 0;
    }
    wprintf(L"\nEmf2EmfPlus - EMF to EMF+ Converter (v1.0)\n");
    LPCWSTR inEmfFilePath = argv[1];
    LPCWSTR outEmfFilePath = argv[2];

    ULONG_PTR gdiplusToken;
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);

    ConvertAndSaveEmfToEmfPlus(inEmfFilePath, outEmfFilePath);

    GdiplusShutdown(gdiplusToken);
    return 0;
}
