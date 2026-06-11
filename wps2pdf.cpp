#include <windows.h>
#include <oleauto.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <strsafe.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

// ============================================================
// COM Helpers
// ============================================================

static void VarSetError(VARIANT* v)
{
    v->vt = VT_ERROR;
    v->scode = DISP_E_PARAMNOTFOUND;
}

static void VarSetBstr(VARIANT* v, const wchar_t* s)
{
    v->vt = VT_BSTR;
    v->bstrVal = SysAllocString(s);
}

static void VarSetI4(VARIANT* v, int val)
{
    v->vt = VT_I4;
    v->lVal = val;
}

static void VarSetBool(VARIANT* v, BOOL val)
{
    v->vt = VT_BOOL;
    v->boolVal = val ? VARIANT_TRUE : VARIANT_FALSE;
}

static HRESULT InvokeMethod(IDispatch* obj, const wchar_t* name,
                            VARIANT* args, UINT cArgs, VARIANT* result)
{
    DISPID id;
    HRESULT hr = obj->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1,
                                     LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;
    DISPPARAMS dp = { args, NULL, cArgs, 0 };
    EXCEPINFO ex = { 0 };
    UINT errArg;
    hr = obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                     DISPATCH_METHOD, &dp, result, &ex, &errArg);
    if (FAILED(hr))
    {
        if (ex.bstrDescription && ex.bstrDescription[0])
            fprintf(stderr, "  COM错误: %ls (wCode=%u)", ex.bstrDescription, ex.wCode);
        else
            fprintf(stderr, "  COM错误: wCode=%u, 来源=%ls", ex.wCode,
                     ex.bstrSource ? ex.bstrSource : L"");
        if (ex.bstrDescription) SysFreeString(ex.bstrDescription);
        if (ex.bstrSource) SysFreeString(ex.bstrSource);
        if (ex.bstrHelpFile) SysFreeString(ex.bstrHelpFile);
    }
    return hr;
}

static HRESULT GetDispatchProp(IDispatch* obj, const wchar_t* name,
                                IDispatch** out)
{
    DISPID id;
    HRESULT hr = obj->GetIDsOfNames(IID_NULL, (LPOLESTR*)&name, 1,
                                     LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;
    VARIANT v;
    VariantInit(&v);
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    hr = obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                     DISPATCH_PROPERTYGET, &dp, &v, NULL, NULL);
    if (SUCCEEDED(hr) && v.vt == VT_DISPATCH && v.pdispVal)
    {
        *out = v.pdispVal;
        (*out)->AddRef();
    }
    else
    {
        hr = E_FAIL;
    }
    VariantClear(&v);
    return hr;
}

// ============================================================
// Architecture detection
// ============================================================

static BOOL IsSystem64Bit()
{
    SYSTEM_INFO si;
    GetNativeSystemInfo(&si);
    return si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64
        || si.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_ARM64;
}

// CLSIDs for 32-bit COM registration (checked via Wow6432Node on 64-bit OS)
static const wchar_t* kWpsProgId    = L"KWps.Application";
static const wchar_t* kWpsClsid     = L"{000209FF-0000-4b30-A977-D214852036FF}";
static const wchar_t* kWordProgId   = L"Word.Application";
static const wchar_t* kWordClsid    = L"{000209FF-0000-0000-C000-000000000046}";

static BOOL IsClsidRegisteredIn64bit(const wchar_t* clsid)
{
    wchar_t key[280];
    StringCchCopyW(key, 280, L"SOFTWARE\\Classes\\CLSID\\");
    StringCchCatW(key, 280, clsid);
    HKEY hKey;
    LONG ret = RegOpenKeyExW(HKEY_LOCAL_MACHINE, key,
        0, KEY_READ | KEY_WOW64_64KEY, &hKey);
    if (ret == ERROR_SUCCESS)
    {
        RegCloseKey(hKey);
        return TRUE;
    }
    return FALSE;
}

// ============================================================
// Help
// ============================================================

static void PrintUsage()
{
    printf("wps2pdf - 将WPS/Office文档(DOC/DOCX)转换为PDF\n");
    printf("Copyright (c) 2025\n");
    printf("\n");
    printf("用法: wps2pdf <输入文件> [-o 输出文件] [-f 起始页] [-t 结束页]\n");
    printf("\n");
    printf("  -o  输出PDF文件路径 (默认: 输入文件同名.pdf)\n");
    printf("  -f  起始页码 (默认: 第1页)\n");
    printf("  -t  结束页码 (默认: 最后一页; 需同时指定 -f)\n");
    printf("\n");
    printf("示例:\n");
    printf("  wps2pdf 1.docx               导出全部页面\n");
    printf("  wps2pdf 1.docx -o out.pdf    指定输出路径\n");
    printf("  wps2pdf 1.docx -f 2 -t 5     导出第2-5页\n");
    printf("  wps2pdf 1.docx -f 3          从第3页导出到末尾\n");
}

// ============================================================
// Entry point
// ============================================================

int wmain(int argc, wchar_t* argv[])
{
    // ---- Parse args ----
    const wchar_t* inputPath = NULL;
    const wchar_t* outputPath = NULL;
    BOOL hasFrom = FALSE, hasTo = FALSE;
    int fromPage = 1, toPage = 1;

    for (int i = 1; i < argc; i++)
    {
        if (wcscmp(argv[i], L"-h") == 0 || wcscmp(argv[i], L"--help") == 0)
        {
            PrintUsage();
            return 0;
        }
        else if ((wcscmp(argv[i], L"-o") == 0 || wcscmp(argv[i], L"--output") == 0) && i + 1 < argc)
        {
            outputPath = argv[++i];
        }
        else if ((wcscmp(argv[i], L"-f") == 0 || wcscmp(argv[i], L"--from") == 0) && i + 1 < argc)
        {
            fromPage = _wtoi(argv[++i]);
            hasFrom = TRUE;
        }
        else if ((wcscmp(argv[i], L"-t") == 0 || wcscmp(argv[i], L"--to") == 0) && i + 1 < argc)
        {
            toPage = _wtoi(argv[++i]);
            hasTo = TRUE;
        }
        else if (argv[i][0] != L'-')
        {
            if (inputPath == NULL)
                inputPath = argv[i];
            else if (outputPath == NULL)
                outputPath = argv[i];
        }
    }

    if (!inputPath)
    {
        PrintUsage();
        return 1;
    }

    if (hasFrom && fromPage < 1)
    {
        fprintf(stderr, "错误: 起始页码必须大于0\n");
        return 1;
    }
    if (hasTo && toPage < 1)
    {
        fprintf(stderr, "错误: 结束页码必须大于0\n");
        return 1;
    }
    if (hasFrom && hasTo && fromPage > toPage)
    {
        fprintf(stderr, "错误: 起始页码不能大于结束页码\n");
        return 1;
    }

    // Auto-generate output path if not specified
    wchar_t autoOutput[MAX_PATH];
    if (!outputPath)
    {
        wcscpy_s(autoOutput, MAX_PATH, inputPath);
        wchar_t* dot = wcsrchr(autoOutput, L'.');
        if (dot)
            wcscpy_s(dot, MAX_PATH - (dot - autoOutput), L".pdf");
        else
            wcscat_s(autoOutput, MAX_PATH, L".pdf");
        outputPath = autoOutput;
    }

    // ---- Initialize COM ----
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
    {
        fprintf(stderr, "错误: COM初始化失败 (0x%08X)\n", hr);
        return 2;
    }

    // ---- Create application (try WPS first, then Office) ----
    static const struct {
        const wchar_t* progId;
        const wchar_t* clsid;
        const wchar_t* displayName;
    } kApps[] = {
        { kWpsProgId,  kWpsClsid,  L"WPS Office" },
        { kWordProgId, kWordClsid, L"Microsoft Office Word" },
    };

    IDispatch* app = NULL;
    const wchar_t* appName = NULL;
    int appIndex = -1;

    for (int i = 0; i < (sizeof(kApps) / sizeof(kApps[0])); i++)
    {
        CLSID clsid;
        hr = CLSIDFromProgID(kApps[i].progId, &clsid);
        if (FAILED(hr)) continue;

        hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER,
                              IID_IDispatch, (void**)&app);
        if (SUCCEEDED(hr))
        {
            appName = kApps[i].displayName;
            appIndex = i;
            break;
        }
    }

    if (FAILED(hr))
    {
        // Try to give specific guidance about 64-bit mismatch
        BOOL found64bit = FALSE;
        for (int i = 0; i < (sizeof(kApps) / sizeof(kApps[0])); i++)
        {
            if (IsSystem64Bit() && IsClsidRegisteredIn64bit(kApps[i].clsid))
            {
                fprintf(stderr, "错误: 检测到64位%ls，但本工具为32位。\n", kApps[i].displayName);
                fprintf(stderr, "请安装32位版本，或使用64位版本的wps2pdf。\n");
                found64bit = TRUE;
                break;
            }
        }
        if (!found64bit)
        {
            fprintf(stderr, "错误: 未找到WPS Office或Microsoft Office。请确保已安装其中之一。\n");
        }
        CoUninitialize();
        return 3;
    }

    printf("检测到: %ls\n", appName);

    // ---- Get Documents collection ----
    IDispatch* docs = NULL;
    hr = GetDispatchProp(app, L"Documents", &docs);
    if (FAILED(hr))
    {
        fprintf(stderr, "错误: 无法获取Documents接口\n");
        app->Release();
        CoUninitialize();
        return 5;
    }

    // ---- Open document ----
    // Documents.Open(FileName, ..., Visible)
    // 15 total params; we pass FileName(1) + fill up to Visible(12)
    // rgvarg[0]=Visible, rgvarg[1..10]=defaults, rgvarg[11]=FileName
    VARIANT openArgs[12];
    for (int i = 0; i < 12; i++) VariantInit(&openArgs[i]);

    openArgs[11].vt = VT_BSTR;    // FileName
    openArgs[11].bstrVal = SysAllocString(inputPath);

    for (int i = 1; i < 11; i++)  // Params 2-11: use defaults
        VarSetError(&openArgs[i]);

    openArgs[0].vt = VT_BOOL;     // Visible = false
    openArgs[0].boolVal = VARIANT_FALSE;

    VARIANT openResult;
    VariantInit(&openResult);
    hr = InvokeMethod(docs, L"Open", openArgs, 12, &openResult);
    docs->Release();

    if (FAILED(hr) || openResult.vt != VT_DISPATCH)
    {
        fprintf(stderr, "错误: 无法打开文档 '%ls'\n", inputPath);
        for (int i = 0; i < 12; i++) VariantClear(&openArgs[i]);
        app->Release();
        CoUninitialize();
        return 6;
    }
    IDispatch* doc = openResult.pdispVal;
    doc->AddRef();
    VariantClear(&openResult);
    for (int i = 0; i < 12; i++) VariantClear(&openArgs[i]);

    // ---- Export to PDF ----
    // ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport,
    //                     OptimizeFor, Range, From, To, Item,
    //                     IncludeDocProps, KeepIRM, CreateBookmarks,
    //                     DocStructureTags, BitmapMissingFonts, UseISO19005_1)
    // Pass only the params we need for maximum compatibility.
    BOOL useRange = hasFrom;
    if (useRange && !hasTo)
        toPage = 9999;  // "to end of document"

    printf("正在转换 [%ls]\n      -> [%ls]\n", inputPath, outputPath);

    if (useRange)
    {
        // Need 7 params: OutputFileName(1) + ExportFormat(2) + Range(5) + From(6) + To(7)
        VARIANT exportArgs[7];
        for (int i = 0; i < 7; i++) VariantInit(&exportArgs[i]);

        VarSetBstr(&exportArgs[6], outputPath);             // OutputFileName
        VarSetI4(&exportArgs[5], 17);                       // ExportFormat=PDF
        VarSetError(&exportArgs[4]);                        // OpenAfterExport (default)
        VarSetError(&exportArgs[3]);                        // OptimizeFor (default)
        VarSetI4(&exportArgs[2], 3);                        // Range=wdExportFromTo
        VarSetI4(&exportArgs[1], fromPage);                 // From
        VarSetI4(&exportArgs[0], toPage);                   // To

        hr = InvokeMethod(doc, L"ExportAsFixedFormat", exportArgs, 7, NULL);
        for (int i = 0; i < 7; i++) VariantClear(&exportArgs[i]);
    }
    else
    {
        // Just 2 required params, rest use defaults
        VARIANT exportArgs[2];
        VariantInit(&exportArgs[0]);
        VariantInit(&exportArgs[1]);

        VarSetBstr(&exportArgs[1], outputPath);             // OutputFileName
        VarSetI4(&exportArgs[0], 17);                       // ExportFormat=PDF

        hr = InvokeMethod(doc, L"ExportAsFixedFormat", exportArgs, 2, NULL);
        VariantClear(&exportArgs[0]);
        VariantClear(&exportArgs[1]);
    }

    if (FAILED(hr))
        fprintf(stderr, "\n错误: PDF导出失败 (0x%08X)\n", hr);
    else
        printf("转换成功\n");

    // ---- Cleanup ----
    InvokeMethod(doc, L"Close", NULL, 0, NULL);
    doc->Release();

    InvokeMethod(app, L"Quit", NULL, 0, NULL);
    app->Release();
    CoUninitialize();

    return SUCCEEDED(hr) ? 0 : 7;
}
