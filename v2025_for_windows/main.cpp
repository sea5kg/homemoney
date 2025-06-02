//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

// #include <vcl\excel.h>

#include <map>

#include "ExcelApp.h"
#include "main.h"
#include "winuser.h"
#include "classeditor.h"
#include "selectclass.h"
#include "Registry.hpp"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
	m_sUnknownClass = L"неизвестно";
	m_strDecDelim = UnicodeString(FormatSettings.DecimalSeparator);
	m_strNumberFormat = L"#,##0.00\"р.\"";
	m_bUseNumberFormat = true;
}
//---------------------------------------------------------------------------

bool TForm1::MakeBackup() {
	UnicodeString dir = ExtractFileDir(m_strFileName) + "\\backups";
	if (!DirectoryExists(dir) && !CreateDir(dir)) {
		return false;
	}
	UnicodeString newfilename = ExtractFileName(m_strFileName);
	newfilename.Delete(newfilename.Length() - ExtractFileExt(m_strFileName).Length() + 1, ExtractFileExt(m_strFileName).Length());

	newfilename =
		dir + L"\\" +
		newfilename + " (" + TDateTime::CurrentDateTime().FormatString(L"yyyy-mm-dd hh_nn") + ")" +
		ExtractFileExt(m_strFileName);
	Log->Lines->Add(L"Сохраняю резервную копию файла в " + newfilename);
	bool bResult = CopyFile(m_strFileName.c_str(),newfilename.c_str(),false);
	return bResult;
}

//---------------------------------------------------------------------------

void TForm1::ReadClassifications(Variant &vSheet, std::vector<exlClass> &classes)
{
	int nRowsCount = vSheet.OlePropertyGet("Cells").OlePropertyGet("Rows").OlePropertyGet("Count");
	int nLastRow1 = vSheet.OlePropertyGet("Cells", nRowsCount, 1).OlePropertyGet("End", xlUp).OlePropertyGet("Row");
	int nLastRow2 = vSheet.OlePropertyGet("Cells", nRowsCount, 2).OlePropertyGet("End", xlUp).OlePropertyGet("Row");
	int nLastRow3 = vSheet.OlePropertyGet("Cells", nRowsCount, 3).OlePropertyGet("End", xlUp).OlePropertyGet("Row");
	int nLastRow4 = vSheet.OlePropertyGet("Cells", nRowsCount, 4).OlePropertyGet("End", xlUp).OlePropertyGet("Row");

	int nLastRow = 0;
	nLastRow = std::max(nLastRow, nLastRow1);
	nLastRow = std::max(nLastRow, nLastRow2);
	nLastRow = std::max(nLastRow, nLastRow3);
	nLastRow = std::max(nLastRow, nLastRow4);
	lblStatus->Caption = L"Считываю классификации...";
	Log->Lines->Add(L"Считываю классификации  всего строк: " + IntToStr(nLastRow-1));
	ProgressBar1->Max = nLastRow;
	ProgressBar1->Min = 0;
	for (int i = 1; i < nLastRow; i++) {
		ProgressBar1->Position = i;
		Application->ProcessMessages();

		exlClass cl;
		cl.Name = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1).OlePropertyGet("Value");
		cl.Class = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,2).OlePropertyGet("Value");
		cl.Comment = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,3).OlePropertyGet("Value");
		cl.Monthes = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OlePropertyGet("Value");

		if (!cl.Name.Trim().IsEmpty()) {
			if (cl.Class.Trim().IsEmpty()) {
			  cl.Class = m_sUnknownClass;
			}
			bool bFound = false;
			for (int s = 0; s < classes.size(); s++) {
				if (classes[s].Name.UpperCase() == cl.Name.UpperCase()) {
					bFound = true;
				}
			}
			if (!bFound) {
				classes.push_back(cl);
			} else {
				Log->Lines->Add(L"При считывании найден дубликат - проигнорирован: [" + cl.Name + "]");
            }
		}
	}
	lblStatus->Caption = L"Готово";
}

//---------------------------------------------------------------------------

void TForm1::WriteClassifications(Variant &vSheet, std::vector<exlClass> &classes)
{
	lblStatus->Caption = L"Приступаю к сортировке классификаций...";
	Log->Lines->Add(L"Приступаю к сортировке классификаций...");
	int nReplaced = 1;

	while(nReplaced > 0) {
		nReplaced = 0;
		for (unsigned int i = 0; i < classes.size()-1; i++) {
		   if (classes[i].Name.UpperCase() > classes[i+1].Name.UpperCase()) {
			   exlClass cl = classes[i];
			   classes[i] = classes[i+1];
			   classes[i+1] = cl;
			   nReplaced++;
		   }
		}
	}
	Log->Lines->Add(L"Отсортировано!");

	int nRowsCount = vSheet.OlePropertyGet("Cells").OlePropertyGet("Rows").OlePropertyGet("Count");
	int nLastRow1 = vSheet.OlePropertyGet("Cells", nRowsCount, 1).OlePropertyGet("End", xlUp).OlePropertyGet("Row");
	int nLastRow2 = vSheet.OlePropertyGet("Cells", nRowsCount, 2).OlePropertyGet("End", xlUp).OlePropertyGet("Row");
	int nLastRow3 = vSheet.OlePropertyGet("Cells", nRowsCount, 3).OlePropertyGet("End", xlUp).OlePropertyGet("Row");
	int nLastRow4 = vSheet.OlePropertyGet("Cells", nRowsCount, 4).OlePropertyGet("End", xlUp).OlePropertyGet("Row");

	int nLastRow = 0;
	nLastRow = std::max(nLastRow, nLastRow1);
	nLastRow = std::max(nLastRow, nLastRow2);
	nLastRow = std::max(nLastRow, nLastRow3);
	nLastRow = std::max(nLastRow, nLastRow4);

	lblStatus->Caption = L"Произвожу очистку классификаций в файле...";
	Log->Lines->Add(L"Произвожу очистку классификаций в файле");
	ProgressBar1->Max = nLastRow;
	ProgressBar1->Min = 0;
	for (int i = 1; i < nLastRow; i++) {
		ProgressBar1->Position = i;
		Application->ProcessMessages();
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1).OlePropertySet("Value", WideString(""));
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,2).OlePropertySet("Value", WideString(""));
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,3).OlePropertySet("Value", WideString(""));
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OlePropertySet("Value", WideString(""));
	}
	lblStatus->Caption = L"Приступаю к записи классификаций в файл...";
	Log->Lines->Add(L"Приступаю к записи в файл " + IntToStr((int)classes.size()));
	ProgressBar1->Max = classes.size();
	ProgressBar1->Min = 0;
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,1).OlePropertySet("Value", WideString(L"Наименование"));
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,2).OlePropertySet("Value", WideString(L"Класс"));
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,3).OlePropertySet("Value", WideString(L"Комментарий"));
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,4).OlePropertySet("Value", WideString(L"Месяца"));
	for (unsigned int i = 0; i < classes.size(); i++) {
		exlClass cl = classes[i];
		ProgressBar1->Position = i;
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+2,1).OlePropertySet("Value", WideString(cl.Name));
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+2,2).OlePropertySet("Value", WideString(cl.Class));
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+2,3).OlePropertySet("Value", WideString(cl.Comment));
		vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+2,4).OlePropertySet("Value", WideString(cl.Monthes));
	}
	lblStatus->Caption = "";
}

//---------------------------------------------------------------------------

void TForm1::ReadMonth(Variant &vSheet, std::vector<ExcelMonthItem> &month)
{
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	int nRowsCount = vSheet.OlePropertyGet("Cells").OlePropertyGet("Rows").OlePropertyGet("Count");
	int nLastRow = vSheet.OlePropertyGet("Cells", nRowsCount, 3).OlePropertyGet("End", xlUp).OlePropertyGet("Row");

	lblStatus->Caption = L"Загрузка данных с листа " + strPageName;
	Log->Lines->Add(L" * Произвожу загрузку данных с листа " + strPageName + L" (строк: " + IntToStr(nLastRow-1) + ")");
	ProgressBar1->Max = nLastRow;
	ProgressBar1->Min = 0;
	int nFound = 0;
	for (int i = 0; i < nLastRow; i++) {
		ProgressBar1->Position = i;
		Application->ProcessMessages();

		ExcelMonthItem mon;
		mon.Month = strPageName;
		mon.Name = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,3).OlePropertyGet("Value");
		if (!mon.Name.Trim().IsEmpty()) {
			UnicodeString value = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OlePropertyGet("Value");
			mon.Price = 0.0f;
			if (value != L"") {
				try {
					mon.Price = value.ToDouble();
				} catch (...) {
					// vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OleFunction("Select", true);
					Log->Lines->Add(L"Ошибка: Не смог преобразовать строку '" + value + L"' в веществененое число" );
					MessageBox (Handle, UnicodeString(L"Ошибка: Не смог преобразовать строку '" + value + L"' в веществененое число").c_str(), L"prompt", MB_OK);
                    return false;
				}
            }
			if (mon.Price != 0) {
				nFound++;
				month.push_back(mon);
			} else {
				Log->Lines->Add(L"\tЧто то не то: " + mon.Name + " : " + FloatToStr(mon.Price));
			}
		}
	}
	ProgressBar1->Position = 0;
	Log->Lines->Add(L" * Найдено записей: " + IntToStr(nFound) + "");
	lblStatus->Caption = "";
}

//---------------------------------------------------------------------------

bool TForm1::ReadMonthSum(Variant &vSheet, double &sum)
{
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	int nRowsCount = vSheet.OlePropertyGet("Cells").OlePropertyGet("Rows").OlePropertyGet("Count");
	int nLastRow = vSheet.OlePropertyGet("Cells", nRowsCount, 4).OlePropertyGet("End", xlUp).OlePropertyGet("Row");

	Log->Lines->Add(L" * Поиск и суммирования сум по дням " + strPageName + L" (строк: " + IntToStr(nLastRow-1) + ")");
	ProgressBar1->Max = nLastRow;
	ProgressBar1->Min = 0;
	int nFound = 0;
	for (int i = 0; i < nLastRow; i++) {
		ProgressBar1->Position = i;
		Application->ProcessMessages();

		ExcelMonthItem mon;
		mon.Month = strPageName;
		mon.Name = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,3).OlePropertyGet("Value");
		if (mon.Name.Trim().IsEmpty()) {
			UnicodeString value = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OlePropertyGet("Value");
			mon.Price = 0.0f;
			if (value != L"") {
				try {
					mon.Price = value.ToDouble();
				} catch (...) {
					// vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OleFunction("Select", true);
					Log->Lines->Add(L"Ошибка: Не смог преобразовать строку '" + value + L"' в веществененое число" );
					MessageBox (Handle, UnicodeString(L"Ошибка: Не смог преобразовать строку '" + value + L"' в веществененое число").c_str(), L"prompt", MB_OK);
                    return false;
				}
            }

			int nWeight = vSheet.
				OlePropertyGet("Cells").
				OlePropertyGet("Item",i+1,4).
				OlePropertyGet("Borders", xlEdgeLeft).
				OlePropertyGet("Weight");
			if (mon.Price != 0 && nWeight == xlMedium) {
				Log->Lines->Add(L"\tСумма: " + mon.Name + ": " + FloatToStr(mon.Price));
				sum += mon.Price;
			}
		}
	}

	ProgressBar1->Position = 0;
	Log->Lines->Add(L" * Найдено: " + IntToStr(nFound) + L", общая сумма: " + FloatToStr(sum));
    return true;
}

//---------------------------------------------------------------------------

void TForm1::setBorders(Variant &vSheet, int nRow, int nCol) {
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);
}

void TForm1::setBordersBold(Variant &vSheet, int nRow, int nCol) {
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertyGet("Borders").OlePropertySet("Weight", xlMedium);
}

void TForm1::clearCell(Variant &vSheet, int nRow, int nCol) {
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertySet("Value", WideString(""));
}

void TForm1::setColor(Variant &vSheet, int nRow, int nCol, int nColor) {
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertyGet("Interior").OlePropertySet("Color", nColor);
}

int TForm1::RGBToInt(int r, int g, int b) {
	TColor cl = (TColor) RGB (r, g, b);
	return cl;
}

void TForm1::ShowErr(const UnicodeString &sErrorMessage) {
	Log->Lines->Add(sErrorMessage);
	MessageBox (Handle, sErrorMessage.c_str(), L"prompt", MB_OK);
	lblStatus->Caption = sErrorMessage;
}

void TForm1::SetSafeFocusOnMainWinow() {
	HWND hWnd = this->Handle;
	bool parentIsVisible = IsWindowVisible(Handle);
	if (this->Enabled && this->Visible && parentIsVisible) {
        this->SetFocus();
	}
}

void TForm1::sort(std::vector<ExcelMonthItem> &months) {
	ProgressBar1->Max = 100;
	ProgressBar1->Min = 0;
	ProgressBar1->Position = 0;

	{
		int nPermutation = 1;
		while (nPermutation > 0) {
			nPermutation = 0;
			for (unsigned int iC = 0; iC < months.size()-1; iC++) {
				ProgressBar1->Position = (ProgressBar1->Position+1) % ProgressBar1->Max;
				Application->ProcessMessages();
				if (months[iC].Class.UpperCase() > months[iC+1].Class.UpperCase()) {
					ExcelMonthItem buf = months[iC];
					months[iC] = months[iC+1];
					months[iC+1] = buf;
					nPermutation++;
				}
			}
		}
		ProgressBar1->Position = 0;
	}
}

//---------------------------------------------------------------------------

UnicodeString TForm1::createHyperLinkToClassification(std::vector<exlClass> &classes, int nLine) {
	UnicodeString name = classes[nLine].Name;
	UnicodeString name2 = "";
	for (int i = 1; i < name.Length()+1; i++) {
		if (name[i] == '"') {
			name2 += "\"\"";
		} else {
			name2 += name[i];
		}
	}

	UnicodeString class_ = classes[nLine].Class;
	UnicodeString class2_ = "";
	for (int i = 1; i < class_.Length()+1; i++) {
		if (class_[i] == '"') {
			class2_ += "\"\"";
		} else {
			class2_ += class_[i];
		}
	}

	UnicodeString sLine = IntToStr(nLine + 2);
	UnicodeString strResult = L"=ГИПЕРССЫЛКА(\"#'классификации'!A" + sLine + ":B" + sLine + L"\", \"" + class2_ + L": " + name2 + L"\")";
	return strResult;
};
//---------------------------------------------------------------------------

void __fastcall TForm1::actCalcClassificationExecute(TObject *Sender)
{
    lblStatus->Caption = L"";
	if(!MakeBackup()) {
		ShowErr(L"Ошибка: Не удалось создать резервную копию файла");
		return;
	}
    ExcelApp app;
	UnicodeString sErrorMessage;
	if (!app.open(m_strFileName, sErrorMessage)) {
		ShowErr(sErrorMessage);
		edtFile->Text = "";
		return;
	}
	app.visible();

	Variant vSheets = app.sheets();
	Variant vSheet = vSheets.OlePropertyGet("Item",m_nPageClassification);
    vSheet.OleFunction("Select", true);
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	if (strPageName.UpperCase() != UnicodeString(L"классификации").UpperCase()) {
	   ShowErr(L"Ошибка: Не верное имя страницы (классификации)");
	   return;
	};

	std::vector<exlClass> classes;
	ReadClassifications(vSheet, classes);

	std::vector<ExcelMonthItem> months;
	Variant vSheetMonth;
	double fSumSum = 0;
	for (unsigned int i = 0; i < m_vMonth.size(); i++) {
		if (cmbMonth->Text == m_vMonth[i].Name) {
			int nMonthPage = m_vMonth[i].Number;
			vSheetMonth = vSheets.OlePropertyGet("Item", nMonthPage);
			vSheetMonth.OleFunction("Select", true);
            SetSafeFocusOnMainWinow();
			ReadMonth(vSheetMonth, months);
			if (!ReadMonthSum(vSheetMonth, fSumSum)) {
                return;
            }
		}
	}

	Log->Lines->Add(L"Произвожу расчет...");
    ProgressBar1->Max = months.size();
	ProgressBar1->Min = 0;
	double nSumHand = 0;

	std::vector<exlSumClass> vSumClasses;
	for (unsigned int i = 0; i < months.size(); i++) {
		ProgressBar1->Position = i;
		int nFound = 0;
		nSumHand += months[i].Price;
		for (unsigned int iC = 0; iC < classes.size(); iC++) {
			if (nFound == 0 && classes[iC].Name.UpperCase() == months[i].Name.UpperCase()) {
				months[i].Class = classes[iC].Class;
				nFound++;
			}
		}
		if (nFound == 0) {
			months[i].Class = m_sUnknownClass;
		}

		nFound = 0;
		for (unsigned int iC = 0; iC < vSumClasses.size(); iC++) {
			if (vSumClasses[iC].Name.UpperCase() == months[i].Class.UpperCase()) {
				vSumClasses[iC].Sum += months[i].Price;
				nFound++;
			}
		}
		if (nFound == 0) {
		   exlSumClass sm;
		   sm.Name = months[i].Class;
		   sm.Sum = months[i].Price;
		   vSumClasses.push_back(sm);
		}
	}

	/*for (unsigned int i = 0; i < vSumClasses.size(); i++) {
		Log->Lines->Add(vSumClasses[i].Name + " = " + FloatToStr(vSumClasses[i].Sum));
	}*/

	Log->Lines->Add(L"Готово.");

//	Log->Lines->Add(L"Сортирую классификации...");



	Log->Lines->Add(L"Сортирую классификации...");
	ProgressBar1->Max = 100;
	ProgressBar1->Min = 0;

	{
		int nPermutation = 1;
		while (nPermutation > 0) {
			nPermutation = 0;
			for (unsigned int iC = 0; iC < vSumClasses.size()-1; iC++) {
				ProgressBar1->Position = (ProgressBar1->Position+1) % ProgressBar1->Max;
				Application->ProcessMessages();
				if (vSumClasses[iC].Name.UpperCase() > vSumClasses[iC+1].Name.UpperCase()) {
					exlSumClass buf = vSumClasses[iC];
					vSumClasses[iC] = vSumClasses[iC+1];
					vSumClasses[iC+1] = buf;
					nPermutation++;
				}
			}
		}
		ProgressBar1->Position = 0;
    }
	Log->Lines->Add(L"Готово.");

	Log->Lines->Add(L"Очистка старых данных");
	ProgressBar1->Max = 100;
	ProgressBar1->Min = 0;

	// clear sum classes, 11,12
	{
		bool b = true;
		int nRow = 1;
		while (b) {
			b = false;
			nRow++;
			ProgressBar1->Position = (ProgressBar1->Position+1) % ProgressBar1->Max;
			Application->ProcessMessages();

			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OleProcedure("ClearFormats");
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OleProcedure("ClearFormats");

			UnicodeString sValue11 = vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertyGet("Value");
			UnicodeString sValue12 = vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertyGet("Value");

			clearCell(vSheetMonth, nRow, 11);
			clearCell(vSheetMonth, nRow, 12);

			if (sValue12.Trim().Length() > 0 || sValue11.Trim().Length() > 0) {
				b = true;
			}
		}
		ProgressBar1->Position = 0;
	}

	// clear 14,15,16,17
	{
		bool b = true;
		int nRow = 1;
		while (b) {
			b = false;
			nRow++;
			ProgressBar1->Position = (ProgressBar1->Position+1) % ProgressBar1->Max;
			Application->ProcessMessages();

			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OleProcedure("ClearFormats");
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OleProcedure("ClearFormats");
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OleProcedure("ClearFormats");
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OleProcedure("ClearFormats");

			UnicodeString sValue14 = vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertyGet("Value");
			UnicodeString sValue15 = vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OlePropertyGet("Value");
			UnicodeString sValue16 = vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertyGet("Value");
			UnicodeString sValue17 = vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertyGet("Value");

			clearCell(vSheetMonth, nRow, 14);
			clearCell(vSheetMonth, nRow, 15);
			clearCell(vSheetMonth, nRow, 16);
			clearCell(vSheetMonth, nRow, 17);

			if (
				sValue14.Trim().Length() > 0
				|| sValue15.Trim().Length() > 0
				|| sValue16.Trim().Length() > 0
				|| sValue17.Trim().Length() > 0
			) {
				b = true;
			}
		}
		ProgressBar1->Position = 0;
	}


	Log->Lines->Add(L"Запись новых данных");
	{
		double nSum = 0;
		int nRow = 2;
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertySet("Value", WideString(L"Класс"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("Value", WideString(L"Сумма"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Columns", WideString("K")).OlePropertySet("ColumnWidth", 20);
		vSheetMonth.OlePropertyGet("Columns", WideString("L")).OlePropertySet("ColumnWidth", 15);
		setBorders(vSheetMonth, nRow, 11);
		setBorders(vSheetMonth, nRow, 12);
		setColor(vSheetMonth, nRow, 11, RGBToInt(240, 230, 140));
		setColor(vSheetMonth, nRow, 12, RGBToInt(240, 230, 140));

		ProgressBar1->Max = vSumClasses.size();
		ProgressBar1->Min = 0;
		for (unsigned int i = 0; i < vSumClasses.size(); i++) {
			nRow++;
			ProgressBar1->Position = i;
			nSum += vSumClasses[i].Sum;
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertySet("Value", WideString(vSumClasses[i].Name));
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("Value", WideString(vSumClasses[i].Sum));
			if (m_bUseNumberFormat) {
				vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
            }

			setBorders(vSheetMonth, nRow, 11);
			setBorders(vSheetMonth, nRow, 12);
		}
		nRow++;
		ProgressBar1->Position = 0;

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertySet("Value", WideString(L"Итого:"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("Value", WideString(nSum));
		if (m_bUseNumberFormat) {
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
        }

		nRow++;

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertySet("Value", WideString(L"Сумма сумм по дням:"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("Value", WideString(fSumSum));
		if (m_bUseNumberFormat) {
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
        }
	}

	Log->Lines->Add(L"Сортирую...");
	sort(months);
    Log->Lines->Add(L"Готово.");

	{
		double nSum = 0;
		int nRow = 2;
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertySet("Value", WideString(L"Класс"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OlePropertySet("Value", WideString(L"Наименование"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("Value", WideString(L"Цена"));

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertyGet("Font").OlePropertySet("Bold", true);

		vSheetMonth.OlePropertyGet("Columns", WideString("N")).OlePropertySet("ColumnWidth", 20);
		vSheetMonth.OlePropertyGet("Columns", WideString("O")).OlePropertySet("ColumnWidth", 50);
		vSheetMonth.OlePropertyGet("Columns", WideString("P")).OlePropertySet("ColumnWidth", 15);

		setBorders(vSheetMonth, nRow, 14);
		setBorders(vSheetMonth, nRow, 15);
		setBorders(vSheetMonth, nRow, 16);

		setColor(vSheetMonth, nRow, 14, RGBToInt(240, 230, 140));
		setColor(vSheetMonth, nRow, 15, RGBToInt(240, 230, 140));
		setColor(vSheetMonth, nRow, 16, RGBToInt(240, 230, 140));

		ProgressBar1->Max = months.size();
		ProgressBar1->Min = 0;
		UnicodeString sLastClass = L"";
		double nSummByClass = 0.0f;
		for (unsigned int i = 0; i < months.size(); i++) {
			nRow++;
			ProgressBar1->Position = i;
			nSum += months[i].Price;

			if (sLastClass != months[i].Class && sLastClass == L"") {
                // first class
				sLastClass = months[i].Class;
			}

			if (sLastClass != months[i].Class && sLastClass != L"") {
			  setBordersBold(vSheetMonth, nRow, 17);
			  vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertyGet("Font").OlePropertySet("Bold", true);
			  vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("Value", WideString(nSummByClass));
			  if (m_bUseNumberFormat) {
				  vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
              }

			  nRow++;
			  sLastClass = months[i].Class;
              nSummByClass = 0.0f;
            }

			nSummByClass += months[i].Price;

			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertySet("Value", WideString(months[i].Class));
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OlePropertySet("Value", WideString(months[i].Name));
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("Value", WideString(months[i].Price));
			if (m_bUseNumberFormat) {
				vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
            }

			if (months[i].Price < 0) {
				setColor(vSheetMonth, nRow, 16, RGBToInt(240, 230, 140));
			}


			// TODO: check linkt to classification
			// vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("Value", WideString(months[i].LinkToClassification.c_str()));

			setBorders(vSheetMonth, nRow, 14);
			setBorders(vSheetMonth, nRow, 15);
			setBorders(vSheetMonth, nRow, 16);
			// setBorders(vSheetMonth, nRow, 17);
		}
		nRow++;
		ProgressBar1->Position = 0;

		setBordersBold(vSheetMonth, nRow, 17);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("Value", WideString(nSummByClass));
		if (m_bUseNumberFormat) {
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
        }

		nRow++;

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertySet("Value", WideString(L"Итого:"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("Value", WideString(nSum));
		if (m_bUseNumberFormat) {
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
		}

	}

	Log->Lines->Add(L"Сохраняю файл...");
	try {
		app.save();
		Log->Lines->Add(L"Данные сохранены!");
	} catch (...) {
		ShowErr(L"Ошибка: Пожалуйста закройте все открытые копии файла и повторите операцию");
		return;
	}
    ProgressBar1->Position = 0;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actCalcClassificationUpdate(TObject *Sender)
{
	actCalcClassification->Enabled = cmbMonth->Text != "";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actViewClassificationsExecute(TObject *Sender)
{
	Variant app = Variant::CreateObject("Excel.Application");
	Variant excel = app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(m_strFileName.c_str()));
	Variant vSheets = excel.OlePropertyGet("Worksheets");
	Variant vSheet = vSheets.OlePropertyGet("Item",m_nPageClassification);
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	if (strPageName.UpperCase() != UnicodeString(L"классификации").UpperCase()) {
	   app.OleProcedure("Quit");
	   ShowErr(L"Ошибка: Не верное имя страницы");
	   return;
	};

	std::vector<exlClass> classes;
	ReadClassifications(vSheet, classes);
	app.OleProcedure("Quit");

	/*RichEdit1->Lines->Clear();
	RichEdit1->Lines->Add("Уровень TreeView1->Selected->Level: "+IntToStr(TreeView1->Selected->Level));
	int n = TreeView1->Selected->AbsoluteIndex;
	RichEdit1->Lines->Add("Асболютный номер TreeView1->Selected->AbsoluteIndex: "+IntToStr(n));
	RichEdit1->Lines->Add("Текст из выбранного узла: "+TreeView1->Selected->Text);*/
	FormSelectClass->ComboBox1->Items->Clear();
	FormClassEditor->TreeView1->Items->Clear();
	std::map<UnicodeString, TTreeNode *> nodeclasses;
	for (unsigned int i = 0; i < classes.size(); i++) {
		if (nodeclasses.count(classes[i].Class) == 0) {
			TTreeNode *parentNode = FormClassEditor->TreeView1->Items->Add(NULL, classes[i].Class);
			nodeclasses[classes[i].Class] = parentNode;
			parentNode->ImageIndex = 0;
			parentNode->SelectedIndex = 0;
			FormSelectClass->ComboBox1->Items->Add(classes[i].Class);
		}
		TTreeNode *pParentNode = nodeclasses[classes[i].Class];
		TTreeNode *pChildNode = FormClassEditor->TreeView1->Items->AddChild(pParentNode, classes[i].Name);
		pChildNode->ImageIndex = 1;
		pChildNode->Data = new exlClass(classes[i]);
		pChildNode->SelectedIndex = 1;
	}


//	TTreeNode *Node1 = FormClassEditor->TreeView1->Items->Add(NULL, "Root");
//	Node1->ImageIndex = 0;
//	FormClassEditor->TreeView1->Items->AddChild(Node1, "Root1");



/*	int n = TreeView1->Selected->AbsoluteIndex;
	TTreeNode *Node1 = TreeView1->Items->Item[n];
	TreeView1->Items->AddChild(Node1,"ChildNode");*/
//	Node1->Selected=true;
    ProgressBar1->Position = 0;
	FormClassEditor->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actViewClassificationsUpdate(TObject *Sender)
{
  actViewClassifications->Enabled = edtFile->Text != "";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actRedesignClassificationsUpdate(TObject *Sender)
{
  actRedesignClassifications->Enabled = edtFile->Text != "";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actRedesignClassificationsExecute(TObject *Sender)
{
	lblStatus->Caption = "";
	if(!MakeBackup()) {
		ShowErr(L"Ошибка: Не удалось создать резервную копию файла");
		return;
	}
	m_strRecomendations = "";

	lblStatus->Caption = L"Открываю файл на чтение...";
	Variant app = Variant::CreateObject("Excel.Application");
    app.OlePropertySet("Visible", true);
	Variant excel = app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(m_strFileName.c_str()));
	Variant vSheets = excel.OlePropertyGet("Worksheets");
	Variant vSheet = vSheets.OlePropertyGet("Item",m_nPageClassification);
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	if (strPageName.UpperCase() != UnicodeString(L"классификации").UpperCase()) {
		app.OleProcedure("Quit");
		ShowErr(L"Ошибка: Не верное имя страницы");
		return;
	};

	std::vector<exlClass> classes;
	ReadClassifications(vSheet, classes);

	std::vector<ExcelMonthItem> months;
	for (unsigned int i = 0; i < m_vMonth.size(); i++) {
		int nMonthPage = m_vMonth[i].Number;
		Variant vSheetMonth = vSheets.OlePropertyGet("Item",nMonthPage);
		ReadMonth(vSheetMonth, months);
	}
//	Log->Lines->Add("Обновляю линки." + IntToStr((int)months.size()));

	int nAll = classes.size() * months.size();
	Log->Lines->Add(L"Всего записей в месяцах: " + IntToStr((int)months.size()));
	std::vector<exlClass> newclasses;
	int nRemovedClasses = 0;
	ProgressBar1->Max = nAll;
	ProgressBar1->Min = 0;
	ProgressBar1->Position = 0;

	lblStatus->Caption = L"Поиск наименований которые отсутвуют в месяцах...";
	Log->Lines->Add(L"Произвожу поиск наименований которые отсутвуют в месяцах...");
	for (unsigned int iC = 0; iC < classes.size(); iC++) {
		exlClass cl = classes[iC];
		cl.Monthes = "";
		UnicodeString name = classes[iC].Name.UpperCase();
		int nCount = 0;
		for (unsigned int iM = 0; iM < months.size(); iM++) {
			ProgressBar1->Position++;
			Application->ProcessMessages();
			UnicodeString name2 = months[iM].Name.UpperCase();
			if (name == name2) {
			   nCount++;
			   if (cl.Monthes.Pos(months[iM].Month) < 1) {
				  cl.Monthes += months[iM].Month + ";";
			   }
			}
		}
		if (nCount > 0) {
			newclasses.push_back(cl);
		} else {
			nRemovedClasses++;
			Log->Lines->Add(L"\tНаименование '" + cl.Name + L"' - нигде не встречается и будет удалено из классификаций");
		}
	}
	Log->Lines->Add(" ** ");

	nAll = newclasses.size() * months.size();
	ProgressBar1->Max = nAll;
	ProgressBar1->Min = 0;
	ProgressBar1->Position = 0;
	int nAddClasses = 0;
	lblStatus->Caption = L"Поиск наименований которые отсутвуют в списке классификаций...";
	Log->Lines->Add(L"Произвожу поиск наименований которые отсутвуют в списке классификаций...");
	for (unsigned int iM = 0; iM < months.size(); iM++) {
		UnicodeString name = months[iM].Name.UpperCase();
		int nCount = 0;
		for (unsigned int iC = 0; iC < newclasses.size(); iC++) {
			ProgressBar1->Position++;
			Application->ProcessMessages();
			UnicodeString name2 = newclasses[iC].Name.UpperCase();
			if (name == name2) {
			   nCount++;
			}
		}
		if (nCount == 0) {
			Log->Lines->Add(L"\tНаименование '" + months[iM].Name + L"' - будет добавлено в классификации");
			exlClass cl;
			cl.Name = months[iM].Name;
			cl.Class = m_sUnknownClass;
			cl.Monthes += months[iM].Month + ";";
			newclasses.push_back(cl);
			nAddClasses++;

			nAll = newclasses.size() * months.size();
			ProgressBar1->Max = nAll;
		}
	}
	Log->Lines->Add(L"Будет удалено классификаций: " + IntToStr(nRemovedClasses));
	Log->Lines->Add(L"Будет добавлено классификаций: " + IntToStr(nAddClasses));
	Log->Lines->Add(L"Всего классификаций: " + IntToStr((int)(newclasses.size())));


	WriteClassifications(vSheet, newclasses);

	lblStatus->Caption = L"Сохраняю файл...";
	Log->Lines->Add(L"Сохраняю файл...");
	try {
		app.OlePropertySet("DisplayAlerts",false);
		excel.OleProcedure("SaveAs", WideString(m_strFileName.c_str()));
		Log->Lines->Add(L"Данные сохранены!");
	} catch (...) {
		ShowErr(L"Ошибка: Пожалуйста закройте все открытые копии файла и повторите операцию");
	}
    ProgressBar1->Position = 0;
	app.OleProcedure("Quit");
	lblStatus->Caption = "";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actSortClassificationsUpdate(TObject *Sender)
{
	actSortClassifications->Enabled = edtFile->Text != "";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actSortClassificationsExecute(TObject *Sender)
{
	if(!MakeBackup()) {
		ShowErr(L"Ошибка: Не удалось создать резервную копию файла");
		return;
	}

	Variant app = Variant::CreateObject("Excel.Application");
	app.OlePropertySet("Visible", true);
	Variant excel = app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(m_strFileName.c_str()));
	Variant vSheets = excel.OlePropertyGet("Worksheets");

	Variant vSheet = vSheets.OlePropertyGet("Item",m_nPageClassification);
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	if (strPageName.UpperCase() != UnicodeString(L"классификации").UpperCase()) {
	   app.OleProcedure("Quit");
	   Log->Lines->Add(L"Ошибка: Не верное имя страницы");
	   MessageBox (Handle, UnicodeString(L"Не верное имя страницы").c_str(), L"prompt", MB_OK);
	   return;
	};
	std::vector<exlClass> classes;
	ReadClassifications(vSheet, classes);
	WriteClassifications(vSheet, classes);


	Log->Lines->Add(L"Сохраняю файл...");
	try {
		app.OlePropertySet("DisplayAlerts",false);
		excel.OleProcedure("SaveAs", WideString(m_strFileName.c_str()));
		Log->Lines->Add(L"Классификации отсортированы!");
	} catch (...) {
		ShowErr(L"Ошибка: Пожалуйста закройте все открытые копии файла и повторите операцию");
	}
	app.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormCreate(TObject *Sender)
{
	lblStatus->Caption = L"";
	TRegistry *reg = new TRegistry();
	reg->RootKey = HKEY_CURRENT_USER;
	reg->OpenKey(L"SOFTWARE\\HomeMoney\\Options", true);
	if (!reg->ValueExists(L"NumberFormat")) {
		reg->WriteString(L"NumberFormat", m_strNumberFormat);
	} else {
		m_strNumberFormat = reg->ReadString(L"NumberFormat");
	}
	if (!reg->ValueExists(L"UseNumberFormat")) {
		reg->WriteBool(L"UseNumberFormat", m_bUseNumberFormat);
	} else {
		m_bUseNumberFormat = reg->ReadBool(L"UseNumberFormat");
	}
	reg->CloseKey();

	delete reg;

	updateLastOpenedFilesList();
		
	menuDisableUseNumberFormat->Enabled = m_bUseNumberFormat;
	menuEnableUseNumberFormat->Enabled = !m_bUseNumberFormat;


//	 TStringList *l=new TStringList;   //Список, в котором будет хранится
//	   TRegistry *reg=new TRegistry();
//	   reg->OpenKey("Software",0);  //Открываем ключ
//	   reg->GetKeyNames(l);
//	   ShowMessage(l->Text);
//	   delete reg;
}
//---------------------------------------------------------------------------

void TForm1::addFileNameToLast(const UnicodeString &strFileName) {

	{
		TRegistry *reg = new TRegistry();
		reg->RootKey = HKEY_CURRENT_USER;
		reg->OpenKey(L"SOFTWARE\\HomeMoney\\LastOpenedFiles", true);
		TStringList *listOfNames = new TStringList();
		reg->GetValueNames(listOfNames);
		TStringList *listOfFiles = new TStringList();
		for (int i = 0; i < listOfNames->Count; i++) {
			listOfFiles->Add(reg->ReadString(listOfNames->Strings[i]));
			reg->DeleteKey(listOfNames->Strings[i]);
		}

		// ShowMessage(L"listOfFiles->Text (1): " + listOfFiles->Text);
				
//		for (int i = 0; i < listOfFiles->Count; i++) {
//			reg->DeleteKey(listOfFiles->Strings[i]);
//		}

		int pos = listOfFiles->IndexOf(strFileName);
		if (pos >= 0) {
			listOfFiles->Delete(pos);
		}
		listOfFiles->Insert(0, strFileName);
		// ShowMessage(L"listOfFiles->Text (2): " + listOfFiles->Text);
	
		for (int i = 0; i < listOfFiles->Count; i++) {
			UnicodeString sName = IntToStr(i);
			while (sName.Length() < 4) {
				sName = L"0" + sName;
            }
			sName = L"File_" + sName;
			reg->WriteString(sName, listOfFiles->Strings[i]);
		}
	
		reg->CloseKey();
	
		delete listOfFiles;
		delete reg;
    }

	updateLastOpenedFilesList();
		
}

void TForm1::updateLastOpenedFilesList() {
	TStringList *listOfFiles = new TStringList();
	{
		TRegistry *reg = new TRegistry();
		reg->OpenKey(L"SOFTWARE\\HomeMoney\\LastOpenedFiles", true);
		TStringList *listOfNames = new TStringList();
		reg->GetValueNames(listOfNames);
		for (int i = 0; i < listOfNames->Count; i++) {
			listOfFiles->Add(reg->ReadString(listOfNames->Strings[i]));
		}
		reg->CloseKey();
		delete reg;
    }
	
	//	ShowMessage(L"listOfFiles->Text (last): " + listOfFiles->Text);
	menuLastOpenedFiles->Enabled = listOfFiles->Count > 0;

	for (int i = menuLastOpenedFiles->Count-1; i >= 0; i--) {
		menuLastOpenedFiles->Delete(i);
	}
	
	for (int i = 0; i < listOfFiles->Count; i++) {
		TMenuItem *pNewItem = new TMenuItem(menuLastOpenedFiles);
		pNewItem->OnClick = clickOpenLastFile;
		pNewItem->Caption = listOfFiles->Strings[i];
		menuLastOpenedFiles->Add(pNewItem);
	}

	delete listOfFiles;

}

void __fastcall TForm1::actOpenExcelFileExecute(TObject *Sender)
{
	if (OpenDialog1->Execute()) {
		openExcelFile(OpenDialog1->FileName);
	}
}
//---------------------------------------------------------------------------

void TForm1::openExcelFile(const UnicodeString &sFileName) {
	edtFile->Text = "";
	m_strFileName = "";

	m_bBackup = false;

	ExcelApp app;
	UnicodeString sErrorMessage;
	if (!app.open(sFileName, sErrorMessage)) {
		MessageBox (Handle, sErrorMessage.c_str(), L"prompt", MB_OK);
		edtFile->Text = "";
		return;
	}

	edtFile->Text = sFileName;
	Log->Lines->Add(L"Файл загружен производиться анализ");
	Variant vSheets = app.sheets();

	m_nPageClassification = 0;
	m_vMonth.clear();
	cmbMonth->Items->Clear();
	cmbMonth->Items->Add("");
	int nSheets = vSheets.OlePropertyGet("Count");
	Log->Lines->Add(L"Всего листов: " + IntToStr(nSheets));
	for (int i = 0; i < nSheets; i++) {
		Variant vSheet = vSheets.OlePropertyGet("Item",i+1);
		UnicodeString str = vSheet.OlePropertyGet("Name");
		// ShowMessage(str);
		if (str.UpperCase() == UnicodeString(L"классификации").UpperCase()) {
			m_nPageClassification = i+1;
		};
		if (str.UpperCase().Pos(L"МЕСЯЦ ") > 0) {
			cmbMonth->Items->Add(str);
			exlSheet s;
			s.Number = i+1;
			s.Name = str;
			m_vMonth.push_back(s);
		}
		Log->Lines->Add(L"Лист " + IntToStr(i+1) + L": " + str);
	}

	if (m_nPageClassification == 0) {
		MessageBox(Handle, UnicodeString(L"Не найден лист 'классификации'").c_str(), L"prompt", MB_OK);
		edtFile->Text = "";
		return;
	}

	if (m_vMonth.size() == 0) {
		MessageBox (Handle, UnicodeString(L"Не найден ни один лист с 'месяц xx'").c_str(), L"prompt", MB_OK);
		edtFile->Text = "";
		return;
	}
	m_strFileName = sFileName;
	addFileNameToLast(sFileName);
}
	
void __fastcall TForm1::menuNumberFormatClick(TObject *Sender)
{
	m_strNumberFormat = InputBox(L"Формат ячеек с суммами", L"Текущий формат", m_strNumberFormat);
	TRegistry *reg = new TRegistry();
	reg->RootKey = HKEY_CURRENT_USER;
	reg->OpenKey(L"SOFTWARE\\HomeMoney\\Options", true);
	reg->WriteString(L"NumberFormat", m_strNumberFormat);
	reg->CloseKey();
	delete reg;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::actUseNumberFormatExecute(TObject *Sender)
{
	m_bUseNumberFormat = !m_bUseNumberFormat;

	TRegistry *reg = new TRegistry();
	reg->RootKey = HKEY_CURRENT_USER;
	reg->OpenKey(L"SOFTWARE\\HomeMoney\\Options", true);
	reg->WriteBool(L"UseNumberFormat", m_bUseNumberFormat);
	reg->CloseKey();
	delete reg;

	menuDisableUseNumberFormat->Enabled = m_bUseNumberFormat;
	menuEnableUseNumberFormat->Enabled = !m_bUseNumberFormat;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::clickOpenLastFile(TObject *Sender)
{
	TMenuItem *pMenuItem = dynamic_cast<TMenuItem *>(Sender);
	if (pMenuItem) {
		UnicodeString filepath = pMenuItem->Caption;
		if (filepath.Pos0(L"&") == 0) {
			filepath = filepath.SubString(2, filepath.Length() - 1);
		}
		openExcelFile(filepath);
//		ShowMessage(L"actOpenLastFileExecute: " + filepath);
	}
}
//---------------------------------------------------------------------------

