
#pragma once

#include <vcl.h>
#pragma hdrstop

class ExcelApp {
    public:
        ExcelApp();
        ~ExcelApp();
		bool open(const UnicodeString &sFileName, UnicodeString &sErrorMessage);
        Variant sheets();

    private:
        bool m_bOpened;
        Variant m_app;
        Variant m_excel;
        Variant m_vSheets;
};