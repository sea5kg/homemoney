
#include "ExcelApp.h"


ExcelApp::ExcelApp() {
	m_bOpened = false;
}

ExcelApp::~ExcelApp() {
	if (m_bOpened) {
		try {
			m_app.OleProcedure("Quit");
		} catch (...) {
		}
	}
}

bool ExcelApp::open(const UnicodeString &sFileName, UnicodeString &sErrorMessage) {
	if (m_bOpened) {
		return true;
	}

	if (!FileExists (sFileName)) {
		sErrorMessage = L"Ошибка: Файл '" + sFileName + L"' не существует";
		return false;
	}

	m_app = Variant::CreateObject("Excel.Application");
	try {
		m_excel = m_app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(sFileName.c_str()));
        // removeFilterDatabase();
	} catch (...) {
		sErrorMessage = L"Ошибка: Не получается открыть файл '" + sFileName + L"' как excel";
		m_app.OleProcedure("Quit");
		// edtFile->Text = "";
		return false;
	}
	m_bOpened = true;
	m_sFileName = sFileName;
	m_vSheets = m_excel.OlePropertyGet("Worksheets");
	return true;
}

void ExcelApp::visible() {
   if (!m_bOpened) {
	   return;
   }
   m_app.OlePropertySet("Visible", true);
}

Variant ExcelApp::sheets() {
	return m_vSheets;
}

void ExcelApp::save() {
    removeFilterDatabase();
	m_app.OlePropertySet("DisplayAlerts",false);
	m_excel.OleProcedure("SaveAs", WideString(m_sFileName.c_str()));
}


void ExcelApp::removeFilterDatabase() {
    Variant vNames = m_excel.OlePropertyGet(L"Names");
	int count = vNames.OlePropertyGet(L"Count");
	std::vector<Variant> vToRemove;
	::OutputDebugStringW(UnicodeString(L"Count => " + UnicodeString(count)).c_str());
	for (int i = 0; i < count; i++) {
		Variant vName = vNames.OleFunction("Item", i+1);
		UnicodeString name = vName.OlePropertyGet("Name");
		if (name.Pos("_FilterDatabase") > 0) {
			::OutputDebugStringW(UnicodeString(UnicodeString(i+1) + L" => " + name).c_str());
			vToRemove.push_back(vName);
//			vNames.OleFunction("Item", i+1).OleProcedure("Delete");
		}

	}

	for (int i = 0; i < vToRemove.size(); i++) {
		vToRemove[i].OleProcedure("Delete");
	}
}

