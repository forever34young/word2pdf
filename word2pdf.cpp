#include <comdef.h>
#include <windows.h>

#include <iostream>
#include <string>
#include <utility>  // for std::pair
#include <vector>

const CLSID CLSID_WordApplication = {
    0x000209FF,
    0x0000,
    0x0000,
    {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
const CLSID CLSID_WpsApplication = {
    0x000209FF,
    0x0000,
    0x4b30,
    {0xA9, 0x77, 0xD2, 0x14, 0x85, 0x20, 0x36, 0xFF}};
const IID IID__Application = {0x00020970,
                              0x0000,
                              0x0000,
                              {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};

DISPPARAMS noArgs = {NULL, NULL, 0, 0};

/**************************************工具函数****************************************************/
double getFloatProp(IDispatch* pDisp, const wchar_t* propName) {
  if (!pDisp) return 0.0;
  DISPID dispid;
  HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&propName, 1,
                                    LOCALE_USER_DEFAULT, &dispid);
  if (FAILED(hr)) return 0.0;

  DISPPARAMS dp = {nullptr, nullptr, 0, 0};
  VARIANT result;
  VariantInit(&result);
  hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                     DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
  if (FAILED(hr)) return 0.0;

  double val = 0.0;
  if (result.vt == VT_R8)
    val = result.dblVal;
  else if (result.vt == VT_R4)
    val = result.fltVal;
  else if (result.vt == VT_I4)
    val = (double)result.lVal;

  VariantClear(&result);
  return val;
}

int getIntProp(IDispatch* pDisp, const wchar_t* propName) {
  return static_cast<int>(getFloatProp(pDisp, propName));
}

bool getProp(IDispatch* pDisp, const wchar_t* propName, VARIANT* pResult) {
  if (!pDisp) return false;
  DISPID dispid;
  HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&propName, 1,
                                    LOCALE_USER_DEFAULT, &dispid);
  if (FAILED(hr)) return false;

  DISPPARAMS dp = {nullptr, nullptr, 0, 0};
  VariantInit(pResult);
  hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                     DISPATCH_PROPERTYGET, &dp, pResult, nullptr, nullptr);
  return SUCCEEDED(hr);
}

bool IsProgIDRegistered(LPCWSTR progID) {
  // 检查 HKCR\<progID>\CLSID 是否存在
  wchar_t keyPath[256];
  wsprintfW(keyPath, L"%s\\CLSID", progID);
  HKEY hKey = nullptr;
  LONG rc = RegOpenKeyExW(HKEY_CLASSES_ROOT, keyPath, 0, KEY_READ, &hKey);
  if (hKey) RegCloseKey(hKey);
  return (rc == ERROR_SUCCESS);
}

HRESULT CreateAppByProgID(LPCWSTR progID, REFIID riid, void** ppv) {
  CLSID clsid;
  HRESULT hr = CLSIDFromProgID(progID, &clsid);
  if (FAILED(hr)) return hr;
  return CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, riid, ppv);
}
/**************************************!工具函数****************************************************/

// 获取页数
int getPages(IDispatch* pWordApp, HRESULT& hr);
// 删除文档中的批注
int deleteComments(IDispatch* pDoc, HRESULT& hr);
// 获取每一节的起始页面、截止页面以及上、下、左、右边距信息
std::vector<std::vector<double>> GetSectionInfo(IDispatch* pDocument,
                                                IDispatch* pWordApp);
// 锁定文档中的所有区域，在转换时，保持原样输出
void lockAllFields(IDispatch* pDoc);
// 保存为 PDF
int saveAsPdf(IDispatch* pDoc, const wchar_t* outputFilePath, int from, int to);
// 获取页眉页脚的范围
double EstimateMaxHeaderFooterHeight(IDispatch* pSection,
                                     bool isFooter = false);

// 将 Word 文档转换为 PDF
int convertWordToPdfByOffice(const wchar_t* inputFilePath,
                             const wchar_t* outputFilePath, int from, int to) {
  std::cout << "inputFilePath=" << _bstr_t(inputFilePath) << std::endl;
  HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
  if (FAILED(hr)) {
    std::cerr << "CoInitializeEx failed: 0x" << std::hex << hr << std::endl;
    return 1;
  }
  std::cout << "CoInitialize success" << std::endl;
  IDispatch* pWordApp = nullptr;

  try {
    bool hasWord = IsProgIDRegistered(L"Word.Application");
    bool hasWPS = IsProgIDRegistered(L"KWPS.Application");

    HRESULT hr = E_FAIL;
    if (hasWPS) {
      CLSID clsid;
      HRESULT hrp = CLSIDFromProgID(L"KWPS.Application", &clsid);
      // 创建 Word 应用程序对象
      hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                            (void**)&pWordApp);
    }

    if (FAILED(hr) && hasWord) {
      std::cout << "use wps failed" << std::endl;
      CLSID clsid;
      HRESULT hro = CLSIDFromProgID(L"Word.Application", &clsid);
      hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                            (void**)&pWordApp);
    }
    if (FAILED(hr)) {
      throw std::runtime_error(
          "1 Unable to create Wps/Word application instance");
    }

    // 获取 Documents 对象
    VARIANT result;
    if (!getProp(pWordApp, L"Documents", &result)) {
      throw std::runtime_error("2 Get Documents Failed");
    }
    IDispatch* pDocuments = result.pdispVal;

    // 打开 Word 文档
    OLECHAR* openMethod = _wcsdup(L"Open");
    DISPID dispidOpen;
    hr = pDocuments->GetIDsOfNames(IID_NULL, &openMethod, 1,
                                   LOCALE_USER_DEFAULT, &dispidOpen);
    if (FAILED(hr)) {
      throw std::runtime_error("4 Get Open Failed");
    }
    std::cout << "Get Open success" << std::endl;
    VARIANT fileName;
    fileName.vt = VT_BSTR;
    fileName.bstrVal = SysAllocString(inputFilePath);

    VARIANT openArgs[1];
    openArgs[0] = fileName;

    DISPPARAMS openParams = {openArgs, NULL, 1, 0};
    hr = pDocuments->Invoke(dispidOpen, IID_NULL, LOCALE_USER_DEFAULT,
                            DISPATCH_METHOD, &openParams, &result, NULL, NULL);
    if (FAILED(hr)) {
      throw std::runtime_error("5 Open Word Failed");
    }
    int pages = getPages(pWordApp, hr);
    std::cout << "Get Pages success" << std::endl;
    /****************************************************************/

    IDispatch* pDoc = result.pdispVal;

    int re = deleteComments(pDoc, hr);

    std::vector<double> distance_vec;
    std::cout << "Delete Comments success" << std::endl;
    auto sectioninfo = GetSectionInfo(pDoc, pWordApp);

    lockAllFields(pDoc);
    re = saveAsPdf(pDoc, outputFilePath, from, to);
    std::cout << "Save as PDF success" << std::endl;
    if (re) {
      std::cout << "6 Save PDF failed" << std::endl;
    } else {
      std::cout << "0 done," << _bstr_t(outputFilePath)
                << ",page_count=" << pages << std::endl;
      for (size_t i = 0; i < sectioninfo.size(); i++) {
        int index = 0;
        std::cout << "end="
                  << sectioninfo[i][0]
                  //<< ",end=" << sectioninfo[i][index++]
                  << ",Top=" << sectioninfo[i][1]
                  << ",Bottom=" << sectioninfo[i][2]
                  << ",Letf=" << sectioninfo[i][3]
                  << ",Right=" << sectioninfo[i][4] << std::endl;
      }
    }

    // 关闭文档
    OLECHAR* closeMethod = _wcsdup(L"Close");
    DISPID dispidClose;
    hr = pDoc->GetIDsOfNames(IID_NULL, &closeMethod, 1, LOCALE_USER_DEFAULT,
                             &dispidClose);
    if (FAILED(hr)) {
      throw std::runtime_error("8 Get CLose Failed");
    }
    // 设置 SaveChanges 参数
    VARIANT saveChanges;
    saveChanges.vt = VT_BOOL;
    saveChanges.boolVal = VARIANT_FALSE;  // 不保存更改

    DISPPARAMS closeParams = {&saveChanges, NULL, 1, 0};
    hr = pDoc->Invoke(dispidClose, IID_NULL, LOCALE_USER_DEFAULT,
                      DISPATCH_METHOD, &closeParams, NULL, NULL, NULL);
    if (FAILED(hr)) {
      _com_error err(hr);
      std::cerr << "Close failed: " << err.ErrorMessage() << std::endl;

      throw std::runtime_error("9 CLose Failed");
    }

    // 退出 Word 应用程序
    OLECHAR* quitMethod = _wcsdup(L"Quit");
    DISPID dispidQuit;
    hr = pWordApp->GetIDsOfNames(IID_NULL, &quitMethod, 1, LOCALE_USER_DEFAULT,
                                 &dispidQuit);
    if (FAILED(hr)) {
      throw std::runtime_error("10 Get Quit Failed ");
    }

    DISPPARAMS quitParams = {NULL, NULL, 0, 0};
    hr = pWordApp->Invoke(dispidQuit, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_METHOD, &quitParams, NULL, NULL, NULL);
    if (FAILED(hr)) {
      throw std::runtime_error("11 Quit Word Failed");
    }
  } catch (const std::exception& e) {
    std::cerr << e.what() << std::endl;
    if (pWordApp) {
      pWordApp->Release();
    }

    CoUninitialize();  // 清理 COM 库
    return 1;
  }

  if (pWordApp) {
    pWordApp->Release();
  }

  CoUninitialize();  // 清理 COM 库
  return 0;
}

// 获取页数
int getPages(IDispatch* pWordApp, HRESULT& hr) {
  int pages = 0;
  VARIANT result1;
  VariantInit(&result1);
  // 获取 ActiveWindow 属性
  if (!getProp(pWordApp, L"ActiveWindow", &result1)) {
    throw std::runtime_error("Invoke ActiveWindow Property Failed");
  }

  IDispatch* pActiveWindow = result1.pdispVal;

  // 获取 ActiveWindow 属性
  if (!getProp(pActiveWindow, L"ActivePane", &result1)) {
    throw std::runtime_error("Invoke ActivePane Property Failed");
  }
  IDispatch* pActivePane = result1.pdispVal;

  // 获取 ActiveWindow 属性
  if (!getProp(pActivePane, L"Pages", &result1)) {
    throw std::runtime_error("Invoke Pages Property Failed");
  }
  IDispatch* pPages = result1.pdispVal;

  // 获取 Pages.Count 属性
  if (!getProp(pPages, L"Count", &result1)) {
    throw std::runtime_error("Invoke Count Property Failed");
  }
  pages = result1.intVal;
  return pages;
}

// 删除文档中的批注
int deleteComments(IDispatch* pDoc, HRESULT& hr) {
  // 删除批注
  // 获取文档的 Comments 集合
  VARIANT commentsResult;
  if (!getProp(pDoc, L"Comments", &commentsResult)) {
    throw std::runtime_error("Failed to invoke Comments property.");
  }
  IDispatch* pComments = commentsResult.pdispVal;

  // 获取 Comments.Item 方法的 DISP ID
  OLECHAR* itemMethod = _wcsdup(L"Item");
  DISPID dispidItem;
  hr = pComments->GetIDsOfNames(IID_NULL, &itemMethod, 1, LOCALE_USER_DEFAULT,
                                &dispidItem);
  if (FAILED(hr)) {
    throw std::runtime_error("Failed to get Item method.");
  }

  // 获取 Comments.Count 属性
  VARIANT countResult;
  if (!getProp(pComments, L"Count", &countResult)) {
    throw std::runtime_error("Failed to invoke Count property.");
  }

  long commentCount = countResult.lVal;

  // 循环删除所有批注
  for (long i = commentCount; i >= 1; i--) {
    VARIANT index;
    index.vt = VT_I4;
    index.lVal = i;

    VARIANT commentResult;
    VariantInit(&commentResult);
    DISPPARAMS indexParam = {&index, NULL, 1, 0};
    hr = pComments->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                           DISPATCH_METHOD, &indexParam, &commentResult, NULL,
                           NULL);
    if (FAILED(hr)) {
      throw std::runtime_error("Failed to get Comment item.");
    }

    IDispatch* pComment = commentResult.pdispVal;

    // 调用 Comment.Delete 方法
    OLECHAR* deleteMethod = _wcsdup(L"Delete");
    DISPID dispidDelete;
    hr = pComment->GetIDsOfNames(IID_NULL, &deleteMethod, 1,
                                 LOCALE_USER_DEFAULT, &dispidDelete);
    if (FAILED(hr)) {
      throw std::runtime_error("Failed to get Delete method.");
    }

    DISPPARAMS deleteParams = {NULL, NULL, 0, 0};
    hr = pComment->Invoke(dispidDelete, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_METHOD, &deleteParams, NULL, NULL, NULL);
    if (FAILED(hr)) {
      throw std::runtime_error("Failed to invoke Delete method.");
    }

    pComment->Release();
  }
  return 0;
}

// 获取每一节的起始页面、截止页面以及上、下、左、右边距信息
std::vector<std::vector<double>> GetSectionInfo(IDispatch* pDocument,
                                                IDispatch* pWordApp) {
  std::vector<std::vector<double>> sectionInfo;
  VARIANT result;
  VariantInit(&result);

  if (!getProp(pDocument, L"Sections", &result)) {
    throw std::runtime_error("Call Sections Failed");
  }
  IDispatch* pSections = result.pdispVal;

  // 获取 Sections.Count 属性
  VARIANT countResult;
  if (!getProp(pSections, L"Count", &countResult)) {
    throw std::runtime_error("Invoke Sections Count Property Failed");
  }

  long sectionCount = countResult.lVal;

  // 遍历每个节
  for (long i = 1; i <= sectionCount; ++i) {
    VARIANT index;
    index.vt = VT_I4;
    index.lVal = i;

    VARIANT sectionResult;
    VariantInit(&sectionResult);
    DISPPARAMS indexParam = {&index, NULL, 1, 0};
    OLECHAR* itemMethod = _wcsdup(L"Item");
    DISPID dispidItem;
    auto hr = pSections->GetIDsOfNames(IID_NULL, &itemMethod, 1,
                                       LOCALE_USER_DEFAULT, &dispidItem);
    if (FAILED(hr)) {
      throw std::runtime_error("Failed to get Section item.");
    }
    hr = pSections->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                           DISPATCH_METHOD, &indexParam, &sectionResult, NULL,
                           NULL);
    if (FAILED(hr)) {
      throw std::runtime_error("Failed to invoke Section item.");
    }
    IDispatch* pSection = sectionResult.pdispVal;

    // 获取节的起始页面
    if (!getProp(pSection, L"Range", &result)) {
      throw std::runtime_error("Call Section Start Range Failed");
    }
    IDispatch* pStartRange = result.pdispVal;

    // 获取起始页面信息
    OLECHAR* startInformationMethod = _wcsdup(L"Information");
    DISPID dispidStartInformation;
    hr = pStartRange->GetIDsOfNames(IID_NULL, &startInformationMethod, 1,
                                    LOCALE_USER_DEFAULT,
                                    &dispidStartInformation);
    if (FAILED(hr)) {
      throw std::runtime_error("Get Start Range Information Failed");
    }

    VARIANT startInfoType;
    startInfoType.vt = VT_I4;
    startInfoType.lVal = 3;  // wdActiveEndPageNumber

    DISPPARAMS dpInfo = {&startInfoType, nullptr, 1, 0};
    hr = pStartRange->Invoke(dispidStartInformation, IID_NULL,
                             LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpInfo,
                             &result, NULL, NULL);
    if (FAILED(hr)) {
      throw std::runtime_error("Call Start Range Information Failed");
    }

    int startPage = result.intVal;

    // 获取 PageSetup 对象
    if (!getProp(pSection, L"PageSetup", &result)) {
      throw std::runtime_error("Call PageSetup Failed");
    }

    IDispatch* pPageSetup = result.pdispVal;

    // 获取页边距属性
    std::vector<double> margins(6);
    OLECHAR* marginNames[6];
    marginNames[0] = _wcsdup(L"TopMargin");
    marginNames[1] = _wcsdup(L"BottomMargin");
    marginNames[2] = _wcsdup(L"LeftMargin");
    marginNames[3] = _wcsdup(L"RightMargin");
    marginNames[4] = _wcsdup(L"HeaderDistance");
    marginNames[5] = _wcsdup(L"FooterDistance");

    double headerHeight = EstimateMaxHeaderFooterHeight(pSection);
    double footerHeight = EstimateMaxHeaderFooterHeight(pSection, true);

    for (int j = 0; j < 6; ++j) {
      DISPID dispidMargin;
      hr = pPageSetup->GetIDsOfNames(IID_NULL, &marginNames[j], 1,
                                     LOCALE_USER_DEFAULT, &dispidMargin);
      if (FAILED(hr)) {
        throw std::runtime_error("Get Margin Property Failed");
      }

      hr = pPageSetup->Invoke(dispidMargin, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &noArgs, &result, NULL,
                              NULL);
      if (FAILED(hr)) {
        throw std::runtime_error("Call Margin Property Failed");
      }
      if (result.vt == VT_R8) {  // double 类型
        margins[j] = result.dblVal;
      } else if (result.vt == VT_R4) {  // float 类型
        margins[j] = static_cast<double>(result.fltVal);
      } else {
        margins[j] = 0.0;
      }
      // 验证值的合理范围
      if (margins[j] > 1000.0 || margins[j] < 0.0) {  // 合理范围
        margins[j] = 0.0;
      }
    }
    double headerTotal = headerHeight + margins[4];
    double footerTotal = footerHeight + margins[5];

    double topMargin = margins[0] > headerTotal ? margins[0] : headerTotal;
    double bottomMargin = margins[1] > footerTotal ? margins[1] : footerTotal;

    std::vector<double> currentSectionInfo = {static_cast<double>(startPage),
                                              topMargin, bottomMargin,
                                              margins[2], margins[3]};

    sectionInfo.push_back(currentSectionInfo);

    // 清理
    pStartRange->Release();
    // pEndRange->Release();
    pPageSetup->Release();
    pSection->Release();
  }

  // 清理
  pSections->Release();

  return sectionInfo;
}

// 锁定文档中的所有区域，在转换时，保持原样输出
void lockAllFields(IDispatch* pDoc) {
  HRESULT hr;
  IDispatch* pFields = nullptr;

  // Step 1: 获取 Document.Fields
  {
    VARIANT result;
    if (!getProp(pDoc, L"Fields", &result)) {
      throw std::runtime_error("Failed to get Fields object.");
    }
    pFields = result.pdispVal;
  }

  // Step 2: 获取 Fields.Count
  long count = 0;
  {
    VARIANT result;
    if (!getProp(pFields, L"Count", &result)) {
      throw std::runtime_error("Failed to get Fields.Count.");
    }
    count = result.lVal;
  }

  // Step 3: 遍历所有 Field 并设置 Locked = true
  for (long i = 1; i <= count; ++i) {
    VARIANT index;
    index.vt = VT_I4;
    index.lVal = i;

    IDispatch* pField = nullptr;
    {
      DISPID dispidItem;
      OLECHAR* name = _wcsdup(L"Item");
      hr = pFields->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT,
                                  &dispidItem);
      if (FAILED(hr)) throw std::runtime_error("Failed to get Fields.Item.");

      DISPPARAMS args = {&index, nullptr, 1, 0};
      VARIANT result;
      VariantInit(&result);
      hr = pFields->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                           DISPATCH_METHOD, &args, &result, nullptr, nullptr);
      if (FAILED(hr)) continue;

      pField = result.pdispVal;
    }

    // 设置 Field.Locked = true
    {
      DISPID dispidLocked;
      OLECHAR* name = _wcsdup(L"Locked");
      hr = pField->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT,
                                 &dispidLocked);
      if (FAILED(hr)) {
        pField->Release();
        continue;
      }

      VARIANT val;
      VariantInit(&val);
      val.vt = VT_BOOL;
      val.boolVal = VARIANT_TRUE;

      DISPID putID = DISPID_PROPERTYPUT;
      DISPPARAMS params = {&val, &putID, 1, 1};
      hr = pField->Invoke(dispidLocked, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_PROPERTYPUT, &params, nullptr, nullptr,
                          nullptr);
    }

    pField->Release();
  }

  if (pFields) pFields->Release();
}

// 保存为 PDF
int saveAsPdf(IDispatch* pDoc, const wchar_t* outputFilePath, int from,
              int to) {
  HRESULT hr;
  try {
    OLECHAR* saveAsMethod = _wcsdup(L"ExportAsFixedFormat");
    DISPID dispidSaveAs;
    hr = pDoc->GetIDsOfNames(IID_NULL, &saveAsMethod, 1, LOCALE_USER_DEFAULT,
                             &dispidSaveAs);
    if (FAILED(hr)) {
      throw std::runtime_error("Failed to get ExportAsFixedFormat method.");
    }

    if (from == -1 && to == -1) {
      VARIANT saveArgs[4];

      saveArgs[3].vt = VT_BSTR;
      saveArgs[3].bstrVal = SysAllocString(outputFilePath);

      saveArgs[2].vt = VT_I4;
      saveArgs[2].lVal = 17;  // wdExportFormatPDF

      saveArgs[1].vt = VT_BOOL;
      saveArgs[1].boolVal = VARIANT_FALSE;  // 不打开导出文件

      saveArgs[0].vt = VT_I4;
      saveArgs[0].lVal = 0;  //  wdExportOptimizeForPrint

      DISPPARAMS saveParams = {saveArgs, NULL, 4, 0};

      hr = pDoc->Invoke(dispidSaveAs, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &saveParams, NULL, NULL, NULL);
    } else {
      VARIANT saveArgs[7];
      int num = 6;
      saveArgs[num].vt = VT_BSTR;
      saveArgs[num--].bstrVal = SysAllocString(outputFilePath);

      saveArgs[num].vt = VT_I4;
      saveArgs[num--].lVal = 17;  // wdExportFormatPDF

      saveArgs[num].vt = VT_BOOL;
      saveArgs[num--].boolVal = VARIANT_FALSE;  // 不打开导出文件

      saveArgs[num].vt = VT_I4;
      saveArgs[num--].lVal = 0;  // wdExportOptimizeForPrint

      saveArgs[num].vt = VT_I4;
      saveArgs[num--].lVal = 3;  // Range
      saveArgs[num].vt = VT_I4;
      saveArgs[num--].lVal = from;  // Start page

      saveArgs[num].vt = VT_I4;
      saveArgs[num--].lVal = to;  // End page

      DISPPARAMS saveParams = {saveArgs, NULL, 7, 0};
      hr = pDoc->Invoke(dispidSaveAs, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &saveParams, NULL, NULL, NULL);
    }
    if (FAILED(hr)) {
      throw std::runtime_error("Failed to save as PDF.");
    }
  } catch (const std::exception& e) {
    std::cerr << e.what() << std::endl;
    return -1;
  }
  return 0;
}

#ifndef DISPID_ITEM
#define DISPID_ITEM 0
#endif

// ======== 段落辅助 ========
double GetMaxFontSizeFromParagraph(IDispatch* pPara) {
  if (!pPara) return 0.0;
  VARIANT rangeVar;
  if (!getProp(pPara, L"Range", &rangeVar) || rangeVar.vt != VT_DISPATCH)
    return 0.0;
  IDispatch* pRange = rangeVar.pdispVal;

  VARIANT charsVar;
  if (!getProp(pRange, L"Characters", &charsVar) || charsVar.vt != VT_DISPATCH)
    return 0.0;

  IDispatch* pChars = charsVar.pdispVal;
  long count = getIntProp(pChars, L"Count");
  if (count == 0) {
    pChars->Release();
    return 0.0;
  }

  double max_font_size = 0.0;

  for (long i = 1; i <= count; ++i) {
    VARIANT idx;
    idx.vt = VT_I4;
    idx.lVal = i;
    VARIANT charVar;
    DISPPARAMS dp = {&idx, nullptr, 1, 0};
    if (FAILED(pChars->Invoke(DISPID_ITEM, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_METHOD, &dp, &charVar, nullptr,
                              nullptr)))
      continue;
    if (charVar.vt != VT_DISPATCH) continue;

    IDispatch* pChar = charVar.pdispVal;
    VARIANT fontVar;
    if (!getProp(pChar, L"Font", &fontVar) || fontVar.vt != VT_DISPATCH) {
      pChar->Release();
      continue;
    }

    double size = getFloatProp(fontVar.pdispVal, L"Size");
    if (size > 0 && size < 500) {
      max_font_size = max_font_size > size ? max_font_size : size;
    }

    fontVar.pdispVal->Release();
    pChar->Release();
  }

  pChars->Release();
  return max_font_size;
}

double EstimateParagraphHeight(IDispatch* pPara) {
  if (!pPara) return 0.0;

  int rule = getIntProp(pPara, L"LineSpacingRule");
  double spacing = getFloatProp(pPara, L"LineSpacing");
  double fontSize = GetMaxFontSizeFromParagraph(pPara);

  double lineHeight = 0.0;
  switch (rule) {
    case 0:
      lineHeight = fontSize * 1.0;
      break;
    case 1:
      lineHeight = fontSize * 1.5;
      break;
    case 2:
      lineHeight = fontSize * 2.0;
      break;
    case 3:
      lineHeight = spacing > fontSize ? spacing : fontSize;
      break;
    case 4:
      lineHeight = spacing;
      break;
    case 5:
      lineHeight = spacing;
      break;
    default:
      lineHeight = fontSize * 1.2;
      break;
  }

  double before = getFloatProp(pPara, L"SpaceBefore");
  double after = getFloatProp(pPara, L"SpaceAfter");

  return before + lineHeight + after;
}

// 获取页眉页脚的范围
double EstimateMaxHeaderFooterHeight(IDispatch* pSection, bool isFooter) {
  if (!pSection) return 0.0;

  HRESULT hr;
  double totalHeight = 0.0;

  // 获取 Footers 集合
  VARIANT footerVar;
  if (isFooter) {
    if (!getProp(pSection, L"Footers", &footerVar) ||
        footerVar.vt != VT_DISPATCH)
      return 0.0;
  } else {
    if (!getProp(pSection, L"Headers", &footerVar) ||
        footerVar.vt != VT_DISPATCH)
      return 0.0;
  }
  IDispatch* pFooters = footerVar.pdispVal;

  // Footers(wdHeaderFooterPrimary) → index = 1
  VARIANT idx;
  idx.vt = VT_I4;
  idx.lVal = 1;

  DISPPARAMS dpGetFooter = {&idx, NULL, 1, 0};
  VARIANT footerResult;
  VariantInit(&footerResult);
  hr = pFooters->Invoke(DISPID_ITEM, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &dpGetFooter, &footerResult, NULL,
                        NULL);
  if (FAILED(hr) || footerResult.vt != VT_DISPATCH) return 0.0;
  IDispatch* pFooter = footerResult.pdispVal;

  // 获取 Footer.Range
  VARIANT rngVar;
  if (!getProp(pFooter, L"Range", &rngVar) || rngVar.vt != VT_DISPATCH)
    return 0.0;
  IDispatch* pRange = rngVar.pdispVal;

  double paraHeight = 0.0;
  double shapeHeight = 0.0;

  // 处理段落高度
  VARIANT parasVar;
  if (getProp(pRange, L"Paragraphs", &parasVar) && parasVar.vt == VT_DISPATCH) {
    IDispatch* pParas = parasVar.pdispVal;
    long count = getIntProp(pParas, L"Count");
    for (long i = 1; i <= count; ++i) {
      VARIANT idx;
      idx.vt = VT_I4;
      idx.lVal = i;
      VARIANT paraVar;
      DISPPARAMS dp = {&idx, nullptr, 1, 0};
      if (SUCCEEDED(pParas->Invoke(DISPID_ITEM, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_METHOD, &dp, &paraVar, nullptr,
                                   nullptr)) &&
          paraVar.vt == VT_DISPATCH) {
        paraHeight += EstimateParagraphHeight(paraVar.pdispVal);
        paraVar.pdispVal->Release();
      }
    }
    pParas->Release();
  }

  // 处理内联图形高度
  VARIANT shapesVar;
  if (getProp(pRange, L"InlineShapes", &shapesVar) &&
      shapesVar.vt == VT_DISPATCH) {
    IDispatch* pShapes = shapesVar.pdispVal;
    long count = getIntProp(pShapes, L"Count");
    for (long i = 1; i <= count; ++i) {
      VARIANT idx;
      idx.vt = VT_I4;
      idx.lVal = i;
      VARIANT shpVar;
      DISPPARAMS dp = {&idx, nullptr, 1, 0};
      if (SUCCEEDED(pShapes->Invoke(DISPID_ITEM, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_METHOD, &dp, &shpVar, nullptr,
                                    nullptr)) &&
          shpVar.vt == VT_DISPATCH) {
        shapeHeight += getFloatProp(shpVar.pdispVal, L"Height");
        shpVar.pdispVal->Release();
      }
    }
    pShapes->Release();
  }

  return paraHeight > shapeHeight ? paraHeight : shapeHeight;
}

int wmain(int argc, wchar_t* argv[]) {
  int from = -1;
  int to = -1;
#if 1
  if (argc < 3) {
    std::cout << "not enough para" << std::endl;
    return -1;
  } else if (argc == 4) {
    // 取第3个参数
    wchar_t* fourthArg = argv[3];

    // 转换为 int
    wchar_t* end;
    long value = std::wcstol(fourthArg, &end, 10);  // Base 10 conversion

    // 检查是否有无效字符
    if (*end != L'\0') {
      throw std::invalid_argument("Invalid number format");
    }

    // 检查范围
    if (value < INT_MIN || value > INT_MAX) {
      std::cout << "Value out of int range" << std::endl;
    }

    from = static_cast<int>(value);
  } else if (argc == 5) {
    // 取第3个参数
    wchar_t* thirdArg = argv[3];

    // 转换为 int
    wchar_t* end;
    long value = std::wcstol(thirdArg, &end, 10);  // Base 10 conversion

    // 检查是否有无效字符
    if (*end != L'\0') {
      throw std::invalid_argument("Invalid number format");
    }

    // 检查范围
    if (value < INT_MIN || value > INT_MAX) {
      std::cout << "Value out of int range" << std::endl;
    }

    from = static_cast<int>(value);

    // 取第4个参数
    wchar_t* fourthArg = argv[4];

    // 转换为 int
    wchar_t* end4;
    long value4 = std::wcstol(fourthArg, &end4, 10);  // Base 10 conversion

    // 检查是否有无效字符
    if (*end4 != L'\0') {
      throw std::invalid_argument("Invalid number format");
    }

    // 检查范围
    if (value4 < INT_MIN || value4 > INT_MAX) {
      std::cout << "Value out of int range" << std::endl;
    }

    to = static_cast<int>(value4);
  }
  if (from > to) {
    std::cout << "\"from\" need greater than \"to\"" << std::endl;
  }

  const wchar_t* inputFile = argv[1];
  const wchar_t* outputFile = argv[2];
#else
  const wchar_t* inputFile = L"C:/1/22.docx";
  const wchar_t* outputFile = L"C:/1/1temppdf.pdf";
  // from = 1;
  // to = 50;
#endif
  std::cout << "from=" << from << "to = " << to << std::endl;
  int re = convertWordToPdfByOffice(inputFile, outputFile, from, to);

  return 0;
}