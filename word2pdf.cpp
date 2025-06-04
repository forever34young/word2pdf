#include <windows.h>
#include <comdef.h>
#include <iostream>

#include <string>
#include <vector>

using namespace std;
//
// Word Application CLSID 和 IID
const CLSID CLSID_WordApplication = { 0x000209FF, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46} };
const CLSID CLSID_WpsApplication = { 0x000209FF, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46} };
const IID IID__Application = { 0x00020970, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46} };

// 获取页数
int getPages(IDispatch* pWordApp, HRESULT& hr)
{
	int pages = 0;
	VARIANT result1;
	VariantInit(&result1);
	// 获取 ActiveWindow 属性
	OLECHAR* activeWindowProp = _wcsdup(L"ActiveWindow");
	DISPID dispidActiveWindow;
	hr = pWordApp->GetIDsOfNames(IID_NULL, &activeWindowProp, 1, LOCALE_USER_DEFAULT, &dispidActiveWindow);
	if (FAILED(hr)) {
		throw std::runtime_error("Get ActiveWindow Property Failed");
	}
	DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
	hr = pWordApp->Invoke(dispidActiveWindow, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result1, NULL, NULL);
	if (FAILED(hr)) {
		throw std::runtime_error("Invoke ActiveWindow Property Failed");
	}

	IDispatch* pActiveWindow = result1.pdispVal;

	// 获取 ActivePane 属性
	OLECHAR* activePaneProp = _wcsdup(L"ActivePane");
	DISPID dispidActivePane;
	hr = pActiveWindow->GetIDsOfNames(IID_NULL, &activePaneProp, 1, LOCALE_USER_DEFAULT, &dispidActivePane);
	if (FAILED(hr)) {
		throw std::runtime_error("Get ActivePane Property Failed");
	}

	hr = pActiveWindow->Invoke(dispidActivePane, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result1, NULL, NULL);
	if (FAILED(hr)) {
		throw std::runtime_error("Invoke ActivePane Property Failed");
	}

	IDispatch* pActivePane = result1.pdispVal;

	// 获取 Pages 属性
	OLECHAR* pagesProp = _wcsdup(L"Pages");
	DISPID dispidPages;
	hr = pActivePane->GetIDsOfNames(IID_NULL, &pagesProp, 1, LOCALE_USER_DEFAULT, &dispidPages);
	if (FAILED(hr)) {
		throw std::runtime_error("Get Pages Property Failed");
	}

	hr = pActivePane->Invoke(dispidPages, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result1, NULL, NULL);
	if (FAILED(hr)) {
		throw std::runtime_error("Invoke Pages Property Failed");
	}

	IDispatch* pPages = result1.pdispVal;

	// 获取 Pages.Count 属性
	OLECHAR* countProp = _wcsdup(L"Count");
	DISPID dispidCount;
	hr = pPages->GetIDsOfNames(IID_NULL, &countProp, 1, LOCALE_USER_DEFAULT, &dispidCount);
	if (FAILED(hr)) {
		throw std::runtime_error("Get Count Property Failed");
	}

	hr = pPages->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result1, NULL, NULL);
	if (FAILED(hr)) {
		throw std::runtime_error("Invoke Count Property Failed");
	}
	pages = result1.intVal;
	return pages;
}

HRESULT GetDispProperty(IDispatch* pDisp, LPOLESTR name, IDispatch** out)
{
	DISPID dispid;
	HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) return hr;

	DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
	VARIANT result;
	VariantInit(&result);

	hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, nullptr, nullptr);
	if (FAILED(hr)) return hr;

	if (result.vt == VT_DISPATCH) {
		*out = result.pdispVal;
		return S_OK;
	}

	return E_FAIL;
}


int deleteComments(IDispatch* pDoc, HRESULT& hr)
{
	// 删除批注
	// 获取文档的 Comments 集合
	OLECHAR* commentsProperty = _wcsdup(L"Comments");
	DISPID dispidComments;
	hr = pDoc->GetIDsOfNames(IID_NULL, &commentsProperty, 1, LOCALE_USER_DEFAULT, &dispidComments);
	if (FAILED(hr)) {
		throw runtime_error("Failed to get Comments property.");
	}

	VARIANT commentsResult;
	VariantInit(&commentsResult);
	DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
	hr = pDoc->Invoke(dispidComments, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &commentsResult, NULL, NULL);
	if (FAILED(hr)) {
		throw runtime_error("Failed to invoke Comments property.");
	}

	IDispatch* pComments = commentsResult.pdispVal;

	// 获取 Comments.Item 方法的 DISP ID
	OLECHAR* itemMethod = _wcsdup(L"Item");
	DISPID dispidItem;
	hr = pComments->GetIDsOfNames(IID_NULL, &itemMethod, 1, LOCALE_USER_DEFAULT, &dispidItem);
	if (FAILED(hr)) {
		throw runtime_error("Failed to get Item method.");
	}

	// 获取 Comments.Count 属性
	OLECHAR* countProperty = _wcsdup(L"Count");
	DISPID dispidCount;
	hr = pComments->GetIDsOfNames(IID_NULL, &countProperty, 1, LOCALE_USER_DEFAULT, &dispidCount);
	if (FAILED(hr)) {
		throw runtime_error("Failed to get Count property.");
	}

	VARIANT countResult;
	VariantInit(&countResult);
	hr = pComments->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &countResult, NULL, NULL);
	if (FAILED(hr)) {
		throw runtime_error("Failed to invoke Count property.");
	}

	long commentCount = countResult.lVal;

	// 循环删除所有批注
	for (long i = commentCount; i >= 1; i--) {
		VARIANT index;
		index.vt = VT_I4;
		index.lVal = i;

		VARIANT commentResult;
		VariantInit(&commentResult);
		DISPPARAMS indexParam = { &index, NULL, 1, 0 };
		hr = pComments->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &indexParam, &commentResult, NULL, NULL);
		if (FAILED(hr)) {
			throw runtime_error("Failed to get Comment item.");
		}

		IDispatch* pComment = commentResult.pdispVal;

		// 调用 Comment.Delete 方法
		OLECHAR* deleteMethod = _wcsdup(L"Delete");
		DISPID dispidDelete;
		hr = pComment->GetIDsOfNames(IID_NULL, &deleteMethod, 1, LOCALE_USER_DEFAULT, &dispidDelete);
		if (FAILED(hr)) {
			throw runtime_error("Failed to get Delete method.");
		}

		DISPPARAMS deleteParams = { NULL, NULL, 0, 0 };
		hr = pComment->Invoke(dispidDelete, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &deleteParams, NULL, NULL, NULL);
		if (FAILED(hr)) {
			throw runtime_error("Failed to invoke Delete method.");
		}

		pComment->Release();
	}
	return 0;
}

// 保存为 PDF 的线程函数
int saveAsPdf(IDispatch* pDoc, const wchar_t* outputFilePath, int from, int to) {
	HRESULT hr;
	try {
		OLECHAR* saveAsMethod = _wcsdup(L"ExportAsFixedFormat");
		DISPID dispidSaveAs;
		hr = pDoc->GetIDsOfNames(IID_NULL, &saveAsMethod, 1, LOCALE_USER_DEFAULT, &dispidSaveAs);
		if (FAILED(hr)) {
			throw std::runtime_error("Failed to get ExportAsFixedFormat method.");
		}

		if (from == -1 && to == -1)
		{
			VARIANT saveArgs[4];

			saveArgs[3].vt = VT_BSTR;
			saveArgs[3].bstrVal = SysAllocString(outputFilePath);

			saveArgs[2].vt = VT_I4;
			saveArgs[2].lVal = 17; // wdExportFormatPDF

			saveArgs[1].vt = VT_BOOL;
			saveArgs[1].boolVal = VARIANT_FALSE;// 不打开导出文件

			saveArgs[0].vt = VT_I4;
			saveArgs[0].lVal = 0;//  wdExportOptimizeForPrint

			DISPPARAMS saveParams = { saveArgs, NULL, 4, 0 };

			hr = pDoc->Invoke(dispidSaveAs, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &saveParams, NULL, NULL, NULL);
		}
		else
		{
			VARIANT saveArgs[7];
			int num = 6;
			saveArgs[num].vt = VT_BSTR;
			saveArgs[num--].bstrVal = SysAllocString(outputFilePath);

			saveArgs[num].vt = VT_I4;
			saveArgs[num--].lVal = 17; // wdExportFormatPDF

			saveArgs[num].vt = VT_BOOL;
			saveArgs[num--].boolVal = VARIANT_FALSE; // 不打开导出文件

			saveArgs[num].vt = VT_I4;
			saveArgs[num--].lVal = 0; // wdExportOptimizeForPrint

			saveArgs[num].vt = VT_I4;
			saveArgs[num--].lVal = 3; // Range
			saveArgs[num].vt = VT_I4;
			saveArgs[num--].lVal = from; // Start page

			saveArgs[num].vt = VT_I4;
			saveArgs[num--].lVal = to; // End page

			DISPPARAMS saveParams = { saveArgs, NULL, 7, 0 };
			hr = pDoc->Invoke(dispidSaveAs, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &saveParams, NULL, NULL, NULL);
		}
		if (FAILED(hr)) {
			throw std::runtime_error("Failed to save as PDF.");
		}
	}
	catch (const std::exception& e) {
		std::cerr << e.what() << std::endl;
		return -1;
	}
	return 0;
}

void removeSuffix(std::wstring& input, const std::wstring& suffix) {
	// 检查字符串是否以指定后缀结尾
	if (input.size() >= suffix.size() &&
		input.compare(input.size() - suffix.size(), suffix.size(), suffix) == 0) {
		// 去除后缀
		input.erase(input.size() - suffix.size(), suffix.size());
	}
}

#include <utility>  // for std::pair

std::pair<double, double> GetHeaderFooterDistances(IDispatch* pApp, IDispatch* pSection) {
	HRESULT hr;
	VARIANT result;
	VariantInit(&result);

	double header_dis = -1;
	double footer_dis = -1;

	// 获取 ActiveWindow
	OLECHAR* activeWindowName = _wcsdup(L"ActiveWindow");
	DISPID dispidActiveWindow;
	hr = pApp->GetIDsOfNames(IID_NULL, &activeWindowName, 1, LOCALE_USER_DEFAULT, &dispidActiveWindow);
	if (FAILED(hr)) throw std::runtime_error("Failed to get ActiveWindow");

	DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
	hr = pApp->Invoke(dispidActiveWindow, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	if (FAILED(hr)) throw std::runtime_error("Failed to get ActiveWindow object");
	IDispatch* pWindow = result.pdispVal;

	// 获取 View 对象
	OLECHAR* viewName = _wcsdup(L"View");
	DISPID dispidView;
	hr = pWindow->GetIDsOfNames(IID_NULL, &viewName, 1, LOCALE_USER_DEFAULT, &dispidView);
	if (FAILED(hr)) throw std::runtime_error("Failed to get View");

	hr = pWindow->Invoke(dispidView, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	if (FAILED(hr)) throw std::runtime_error("Failed to get View object");
	IDispatch* pView = result.pdispVal;

	// 设置 SeekView = 1 (wdSeekCurrentPageHeader)
	VARIANT seekVal;
	seekVal.vt = VT_I4;
	seekVal.lVal = 1;
	DISPID dispidSeekView;
	OLECHAR* seekName = _wcsdup(L"SeekView");
	hr = pView->GetIDsOfNames(IID_NULL, &seekName, 1, LOCALE_USER_DEFAULT, &dispidSeekView);
	if (FAILED(hr)) throw std::runtime_error("Failed to get SeekView");

	DISPPARAMS seekParams = { &seekVal, nullptr, 1, 0 };
	hr = pView->Invoke(dispidSeekView, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &seekParams, NULL, NULL, NULL);
	if (FAILED(hr)) throw std::runtime_error("Failed to set SeekView");

	// 获取 Selection 对象
	OLECHAR* selName = _wcsdup(L"Selection");
	DISPID dispidSelection;
	hr = pApp->GetIDsOfNames(IID_NULL, &selName, 1, LOCALE_USER_DEFAULT, &dispidSelection);
	if (FAILED(hr)) throw std::runtime_error("Failed to get Selection");

	hr = pApp->Invoke(dispidSelection, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	if (FAILED(hr)) throw std::runtime_error("Failed to get Selection object");
	IDispatch* pSelection = result.pdispVal;

	// 获取 Selection.Range.Information(VerticalPositionRelativeToPage)
	DISPID dispidRange;
	OLECHAR* rangeName = _wcsdup(L"Range");
	hr = pSelection->GetIDsOfNames(IID_NULL, &rangeName, 1, LOCALE_USER_DEFAULT, &dispidRange);
	if (FAILED(hr)) throw std::runtime_error("Failed to get Range");

	hr = pSelection->Invoke(dispidRange, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	if (FAILED(hr)) throw std::runtime_error("Failed to get Range object");
	IDispatch* pHeaderRange = result.pdispVal;

	// 调用 Range.Information(VerticalPositionRelativeToPage)
	OLECHAR* infoName = _wcsdup(L"Information");
	DISPID dispidInfo;
	hr = pHeaderRange->GetIDsOfNames(IID_NULL, &infoName, 1, LOCALE_USER_DEFAULT, &dispidInfo);
	if (FAILED(hr)) throw std::runtime_error("Failed to get Information");

	VARIANT infoType;
	infoType.vt = VT_I4;
	infoType.lVal = 4; // wdVerticalPositionRelativeToPage

	DISPPARAMS dpInfo = { &infoType, nullptr, 1, 0 };
	hr = pHeaderRange->Invoke(dispidInfo, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpInfo, &result, NULL, NULL);
	if (SUCCEEDED(hr)) header_dis = result.dblVal;

	pHeaderRange->Release();

	// 设置 SeekView = 2 (wdSeekCurrentPageFooter)
	seekVal.lVal = 2;
	hr = pView->Invoke(dispidSeekView, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &seekParams, NULL, NULL, NULL);
	if (FAILED(hr)) throw std::runtime_error("Failed to set SeekView to Footer");

	// 重新获取 Selection -> Range
	hr = pApp->Invoke(dispidSelection, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	IDispatch* pSelFooter = result.pdispVal;

	hr = pSelFooter->Invoke(dispidRange, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	IDispatch* pFooterRange = result.pdispVal;

	infoType.lVal = 4; // VerticalPositionRelativeToPage
	hr = pFooterRange->Invoke(dispidInfo, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpInfo, &result, NULL, NULL);
	if (SUCCEEDED(hr)) footer_dis = result.dblVal;

	pFooterRange->Release();
	pSelFooter->Release();
	pSelection->Release();
	pView->Release();
	pWindow->Release();

	return { header_dis, footer_dis };
}


// 获取每一节的起始页面、截止页面以及上、下、左、右边距信息
std::vector<std::vector<double>> GetSectionInfo( IDispatch* pDocument, IDispatch* pWordApp) {
	std::vector<std::vector<double>> sectionInfo;
	VARIANT result;
	VariantInit(&result);

	// 获取 Sections 集合
	OLECHAR* sectionsMethod = _wcsdup(L"Sections");
	DISPID dispidSections;
	HRESULT hr = pDocument->GetIDsOfNames(IID_NULL, &sectionsMethod, 1, LOCALE_USER_DEFAULT, &dispidSections);
	if (FAILED(hr)) {
		throw std::runtime_error("Get Sections Failed");
	}
	DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
	hr = pDocument->Invoke(dispidSections, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	if (FAILED(hr)) {
		throw std::runtime_error("Call Sections Failed");
	}

	IDispatch* pSections = result.pdispVal;

	// 获取 Sections.Count 属性
	OLECHAR* countProperty = _wcsdup(L"Count");
	DISPID dispidCount;
	hr = pSections->GetIDsOfNames(IID_NULL, &countProperty, 1, LOCALE_USER_DEFAULT, &dispidCount);
	if (FAILED(hr)) {
		throw std::runtime_error("Get Sections Count Property Failed");
	}

	VARIANT countResult;
	VariantInit(&countResult);
	hr = pSections->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &countResult, NULL, NULL);
	if (FAILED(hr)) {
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
		DISPPARAMS indexParam = { &index, NULL, 1, 0 };
		OLECHAR* itemMethod = _wcsdup(L"Item");
		DISPID dispidItem;
		hr = pSections->GetIDsOfNames(IID_NULL, &itemMethod, 1, LOCALE_USER_DEFAULT, &dispidItem);
		if (FAILED(hr)) {
			throw std::runtime_error("Failed to get Section item.");
		}
		hr = pSections->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &indexParam, &sectionResult, NULL, NULL);
		if (FAILED(hr)) {
			throw std::runtime_error("Failed to invoke Section item.");
		}

		IDispatch* pSection = sectionResult.pdispVal;

		// 获取节的起始页面
		OLECHAR* startRangeMethod = _wcsdup(L"Range");
		DISPID dispidStartRange;
		hr = pSection->GetIDsOfNames(IID_NULL, &startRangeMethod, 1, LOCALE_USER_DEFAULT, &dispidStartRange);
		if (FAILED(hr)) {
			throw std::runtime_error("Get Section Start Range Failed");
		}

		hr = pSection->Invoke(dispidStartRange, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
		if (FAILED(hr)) {
			throw std::runtime_error("Call Section Start Range Failed");
		}

		IDispatch* pStartRange = result.pdispVal;

		// 获取起始页面信息
		OLECHAR* startInformationMethod = _wcsdup(L"Information");
		DISPID dispidStartInformation;
		hr = pStartRange->GetIDsOfNames(IID_NULL, &startInformationMethod, 1, LOCALE_USER_DEFAULT, &dispidStartInformation);
		if (FAILED(hr)) {
			throw std::runtime_error("Get Start Range Information Failed");
		}

		VARIANT startInfoType;
		startInfoType.vt = VT_I4;
		startInfoType.lVal = 3; // wdActiveEndPageNumber

		DISPPARAMS dpInfo = { &startInfoType, nullptr, 1, 0 };
		hr = pStartRange->Invoke(dispidStartInformation, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpInfo, &result, NULL, NULL);
		if (FAILED(hr)) {
			throw std::runtime_error("Call Start Range Information Failed");
		}

		int startPage = result.intVal;

		// 获取 PageSetup 对象
		OLECHAR* pageSetupMethod = _wcsdup(L"PageSetup");
		DISPID dispidPageSetup;
		hr = pSection->GetIDsOfNames(IID_NULL, &pageSetupMethod, 1, LOCALE_USER_DEFAULT, &dispidPageSetup);
		if (FAILED(hr)) {
			throw std::runtime_error("Get PageSetup Failed");
		}

		hr = pSection->Invoke(dispidPageSetup, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
		if (FAILED(hr)) {
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

		for (int j = 0; j < 6; ++j) {
			DISPID dispidMargin;
			hr = pPageSetup->GetIDsOfNames(IID_NULL, &marginNames[j], 1, LOCALE_USER_DEFAULT, &dispidMargin);
			if (FAILED(hr)) {
				throw std::runtime_error("Get Margin Property Failed");
			}

			hr = pPageSetup->Invoke(dispidMargin, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
			if (FAILED(hr)) {
				throw std::runtime_error("Call Margin Property Failed");
			}
			if (result.vt == VT_R8) { // double 类型
				margins[j] = result.dblVal;
			}
			else if (result.vt == VT_R4) { // float 类型
				margins[j] = static_cast<double>(result.fltVal);
			}
			else {
				margins[j] = 0.0;
			}
			// 验证值的合理范围
			if (margins[j] > 1000.0 || margins[j] < 0.0) { // 合理范围
				margins[j] = 0.0;
			}
		}

		std::vector<double> currentSectionInfo = { static_cast<double>(startPage),  margins[0], margins[1], margins[2], margins[3], margins[4], margins[5] };
		
		sectionInfo.push_back(currentSectionInfo);

		
		// 清理
		pStartRange->Release();
		//pEndRange->Release();
		pPageSetup->Release();
		pSection->Release();
	}

	// 清理
	pSections->Release();

	return sectionInfo;
}

void lockAllFields(IDispatch* pDoc)
{
	HRESULT hr;
	IDispatch* pFields = nullptr;

	// Step 1: 获取 Document.Fields
	{
		DISPID dispidFields;
		OLECHAR* name = _wcsdup(L"Fields");
		hr = pDoc->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispidFields);
		if (FAILED(hr)) throw std::runtime_error("Failed to get Fields property.");

		DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
		VARIANT result;
		VariantInit(&result);
		hr = pDoc->Invoke(dispidFields, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, nullptr, nullptr);
		if (FAILED(hr)) throw std::runtime_error("Failed to get Fields object.");

		pFields = result.pdispVal;
	}

	// Step 2: 获取 Fields.Count
	long count = 0;
	{
		DISPID dispidCount;
		OLECHAR* name = _wcsdup(L"Count");
		hr = pFields->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispidCount);
		if (FAILED(hr)) throw std::runtime_error("Failed to get Count property.");

		VARIANT result;
		VariantInit(&result);
		DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
		hr = pFields->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, nullptr, nullptr);
		if (FAILED(hr)) throw std::runtime_error("Failed to get Fields.Count.");

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
			hr = pFields->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispidItem);
			if (FAILED(hr)) throw std::runtime_error("Failed to get Fields.Item.");

			DISPPARAMS args = { &index, nullptr, 1, 0 };
			VARIANT result;
			VariantInit(&result);
			hr = pFields->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &args, &result, nullptr, nullptr);
			if (FAILED(hr)) continue;

			pField = result.pdispVal;
		}

		// 设置 Field.Locked = true
		{
			DISPID dispidLocked;
			OLECHAR* name = _wcsdup(L"Locked");
			hr = pField->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispidLocked);
			if (FAILED(hr)) {
				pField->Release();
				continue;
			}

			VARIANT val;
			VariantInit(&val);
			val.vt = VT_BOOL;
			val.boolVal = VARIANT_TRUE;

			DISPID putID = DISPID_PROPERTYPUT;
			DISPPARAMS params = { &val, &putID, 1, 1 };
			hr = pField->Invoke(dispidLocked, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, nullptr, nullptr, nullptr);
		}

		pField->Release();
	}

	if (pFields) pFields->Release();
}


// 将 Word 文档转换为 PDF
int convertWordToPdfByOffice(const wchar_t* inputFilePath, const wchar_t* outputFilePath, int from, int to) {
	std::cout << "inputFilePath=" << _bstr_t(inputFilePath) << std::endl;
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		std::cerr << "CoInitializeEx failed: 0x" << std::hex << hr << std::endl;
		return 1;
	}
	std::cout << "CoInitialize success" << std::endl;
	IDispatch* pWordApp = nullptr;

	try {
		std::cout << "try" << std::endl;
		// 创建 Word 应用程序对象
		HRESULT hr = CoCreateInstance(CLSID_WordApplication, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pWordApp);
		if (FAILED(hr)) {
			throw runtime_error("1 Unable to create Word application instance");
		}
		std::cout << "Word application instance created successfully." << std::endl;
		// 获取 Documents 对象
		VARIANT result;
		VariantInit(&result);
		OLECHAR* documentsMethod = _wcsdup(L"Documents");
		DISPID dispidDocuments;

		hr = pWordApp->GetIDsOfNames(IID_NULL, &documentsMethod, 1, LOCALE_USER_DEFAULT, &dispidDocuments);
		if (FAILED(hr)) {
			throw runtime_error("2 Get Documents Failed");
		}
		std::cout << "Get Documents success" << std::endl;
		DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
		hr = pWordApp->Invoke(dispidDocuments, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
		if (FAILED(hr)) {
			throw runtime_error("3 Call Documents Failed");
		}
		std::cout << "Call Documents success" << std::endl;
		IDispatch* pDocuments = result.pdispVal;

		// 打开 Word 文档
		OLECHAR* openMethod = _wcsdup(L"Open");
		DISPID dispidOpen;
		hr = pDocuments->GetIDsOfNames(IID_NULL, &openMethod, 1, LOCALE_USER_DEFAULT, &dispidOpen);
		if (FAILED(hr)) {
			throw runtime_error("4 Get Open Failed");
		}
		std::cout << "Get Open success" << std::endl;
		VARIANT fileName;
		fileName.vt = VT_BSTR;
		fileName.bstrVal = SysAllocString(inputFilePath);

		VARIANT openArgs[1];
		openArgs[0] = fileName;

		DISPPARAMS openParams = { openArgs, NULL, 1, 0 };
		hr = pDocuments->Invoke(dispidOpen, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &openParams, &result, NULL, NULL);
		if (FAILED(hr)) {
			throw runtime_error("5 Open Word Failed");
		}
		int pages = getPages(pWordApp, hr);
		std::cout << "Get Pages success" << std::endl;
		/****************************************************************/

		IDispatch* pDoc = result.pdispVal;

		int re = deleteComments(pDoc, hr);

		std::vector<double> distance_vec;
		std::cout << "Delete Comments success" << std::endl;
		auto sectioninfo = GetSectionInfo(pDoc,pWordApp);

		lockAllFields(pDoc);
		re = saveAsPdf(pDoc, outputFilePath, from, to);
		std::cout << "Save as PDF success" << std::endl;
		if (re)
		{
			cout << "6 failed" << endl;
		}
		else
		{
			cout << "0 done," << _bstr_t(outputFilePath) << ",page_count=" << pages << std::endl;
			for (size_t i = 0; i < sectioninfo.size(); i++)
			{
				int index = 0;
				std::cout
					<< "end=" << sectioninfo[i][0]
					//<< ",end=" << sectioninfo[i][index++]
					<< ",Top=" << sectioninfo[i][1]
					<< ",Bottom=" << sectioninfo[i][2]
					<< ",Letf=" << sectioninfo[i][3]
					<< ",Right=" << sectioninfo[i][4]
					<< ",Header=" << sectioninfo[i][5]
					<< ",Footer=" << sectioninfo[i][6]
					<< endl;
			}
		}

		// 关闭文档
		OLECHAR* closeMethod = _wcsdup(L"Close");
		DISPID dispidClose;
		hr = pDoc->GetIDsOfNames(IID_NULL, &closeMethod, 1, LOCALE_USER_DEFAULT, &dispidClose);
		if (FAILED(hr)) {
			throw runtime_error("8 Get CLose Failed");
		}
		// 设置 SaveChanges 参数
		VARIANT saveChanges;
		saveChanges.vt = VT_BOOL;
		saveChanges.boolVal = VARIANT_FALSE; // 不保存更改

		DISPPARAMS closeParams = { &saveChanges, NULL, 1, 0 };
		hr = pDoc->Invoke(dispidClose, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &closeParams, NULL, NULL, NULL);
		if (FAILED(hr)) {
			_com_error err(hr);
			cerr << "Close failed: " << err.ErrorMessage() << endl;

			throw runtime_error("9 CLose Failed");
		}

		// 退出 Word 应用程序
		OLECHAR* quitMethod = _wcsdup(L"Quit");
		DISPID dispidQuit;
		hr = pWordApp->GetIDsOfNames(IID_NULL, &quitMethod, 1, LOCALE_USER_DEFAULT, &dispidQuit);
		if (FAILED(hr)) {
			throw runtime_error("10 Get Quit Failed ");
		}

		DISPPARAMS quitParams = { NULL, NULL, 0, 0 };
		hr = pWordApp->Invoke(dispidQuit, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &quitParams, NULL, NULL, NULL);
		if (FAILED(hr)) {
			throw runtime_error("11 Quit Word Failed");
		}
	}
	catch (const exception& e) {
		cerr << e.what() << endl;
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

int wmain(int argc, wchar_t* argv[]) {
	int from = -1;
	int to = -1;
#if 1
	if (argc < 3) {
		std::cout << "not enough para" << std::endl;
		return -1;
	}
	else if (argc == 3)
	{
		std::cout << "para 3" << std::endl;
	}
	else if (argc == 4)
	{
		// 取第3个参数
		wchar_t* fourthArg = argv[3];

		// 转换为 int
		wchar_t* end;
		long value = std::wcstol(fourthArg, &end, 10); // Base 10 conversion

		// 检查是否有无效字符
		if (*end != L'\0') {
			throw std::invalid_argument("Invalid number format");
		}

		// 检查范围
		if (value < INT_MIN || value > INT_MAX) {
			std::cout << "Value out of int range" << std::endl;
		}

		from = static_cast<int>(value);
	}
	else if (argc == 5)
	{
		// 取第四个参数
		wchar_t* thirdArg = argv[3];

		// 转换为 int
		wchar_t* end;
		long value = std::wcstol(thirdArg, &end, 10); // Base 10 conversion

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
		long value4 = std::wcstol(fourthArg, &end4, 10); // Base 10 conversion

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
	if (from > to)
	{
		std::cout << "\"from\" need greater than \"to\"" << std::endl;
	}

	const wchar_t* inputFile = argv[1];
	const wchar_t* outputFile = argv[2];
#else
	const wchar_t* inputFile = L"F:/1234.doc";
	const wchar_t* outputFile = L"F:/1234.pdf";
	//from = 1;
	//to = 50;
#endif
	std::cout << "from=" << from << "to = " << to << std::endl;
	int re = convertWordToPdfByOffice(inputFile, outputFile, from, to);
	if (!re)
	{
		return 0;
	}

	//convertWordToPdfByWps(inputFile, outputFile);

	return 0;
}