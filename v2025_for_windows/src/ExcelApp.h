
#pragma once

#include <vcl.h>
#pragma hdrstop

class ExcelApp {
    public:
        ExcelApp();
        ~ExcelApp();
		bool open(const UnicodeString &sFileName, UnicodeString &sErrorMessage);
        void visible();
		Variant sheets();
        void save();

	private:
		void removeFilterDatabase();

		bool m_bOpened;
        UnicodeString m_sFileName;
        Variant m_app;
        Variant m_excel;
        Variant m_vSheets;
};