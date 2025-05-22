
#include "ExcelApp.h"


ExcelApp::ExcelApp() {
    m_bOpened = false;
}

ExcelApp::~ExcelApp() {
}

bool ExcelApp::open(UnicodeString strFileName) {
    if (!m_bOpened) {
        m_app = Variant::CreateObject("Excel.Application");
        try {
            m_excel = m_app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(strFileName.c_str()));
        } catch (...) {
            MessageBox (Handle, UnicodeString(L"Не получается открыть файл '" + strFileName + L"' как excel").c_str(), L"prompt", MB_OK);
            m_app.OleProcedure("Quit");
            // edtFile->Text = "";
            return false;
        }
        m_bOpened = true;
        m_vSheets = excel.OlePropertyGet("Worksheets");
    }
    return true;
}