//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

// #include <vcl\excel.h>

#include <map>

#include "main.h"
#include "classeditor.h"
#include "selectclass.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
	m_sUnknownClass = "неизвестно";
	m_strDecDelim = UnicodeString(FormatSettings.DecimalSeparator);
	m_strNumberFormat = "# ##0" + m_strDecDelim + "00р.";
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button1Click(TObject *Sender)
{
	edtFile->Text = "";
	m_strFileName = "";
	if (OpenDialog1->Execute()) {
		m_bBackup = false;

		UnicodeString strFileName = OpenDialog1->FileName;
		if (! FileExists (strFileName)) {
			MessageBox (Handle, UnicodeString(L"Файл '" + strFileName + L"' не существует").c_str(), L"prompt", MB_OK);
			edtFile->Text = "";
			return;
		}

		edtFile->Text = strFileName;
		// TODO: create backup now or later?
		// TODO: INIT OPERATIONS
		// TODO: scan month
		Variant var_Book,var_Sheet,var_Cell;

		Variant app = Variant::CreateObject("Excel.Application");
		// app.OlePropertySet("Visible",false);
		Variant excel;
		try {
			excel = app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(strFileName.c_str()));
		} catch (...) {
			MessageBox (Handle, UnicodeString(L"Не получается открыть файл '" + strFileName + L"' как excel").c_str(), L"prompt", MB_OK);
			app.OleProcedure("Quit");
			edtFile->Text = "";
			return;
		}

		Log->Lines->Add("Файл загружен производиться анализ");
		Variant vSheets = excel.OlePropertyGet("Worksheets");

		m_nPageClassification = 0;
		m_vMonth.clear();
		cmbMonth->Items->Clear();
		cmbMonth->Items->Add("");
		int nSheets = vSheets.OlePropertyGet("Count");
		Log->Lines->Add("Всего листов: " + IntToStr(nSheets));
		for (int i = 0; i < nSheets; i++) {
			Variant vSheet = vSheets.OlePropertyGet("Item",i+1);
			UnicodeString str = vSheet.OlePropertyGet("Name");
			if (str.UpperCase() == UnicodeString("классификации").UpperCase()) {
				m_nPageClassification = i+1;
			};
			if (str.UpperCase().Pos("МЕСЯЦ ") > 0) {
				cmbMonth->Items->Add(str);
				exlSheet s;
				s.Number = i+1;
				s.Name = str;
				m_vMonth.push_back(s);
			}
			Log->Lines->Add("Лист " + IntToStr(i+1) + ": " + str);
		}

		if (m_nPageClassification == 0) {
			MessageBox (Handle, UnicodeString(L"Не найден лист 'классификации'").c_str(), L"prompt", MB_OK);
			app.OleProcedure("Quit");
			edtFile->Text = "";
			return;
		}

		if (m_vMonth.size() == 0) {
			MessageBox (Handle, UnicodeString(L"Не найден ни один лист с 'месяц xx'").c_str(), L"prompt", MB_OK);
			app.OleProcedure("Quit");
			edtFile->Text = "";
			return;
        }
		app.OleProcedure("Quit");
		m_strFileName = strFileName;
    }
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
		newfilename + " (" + TDateTime::CurrentDateTime().FormatString("yyyy-mm-dd hh_nn") + ")" +
		ExtractFileExt(m_strFileName);
	Log->Lines->Add("Сохраняю резервную копию файла в " + newfilename);
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
	lblStatus->Caption = "Считываю классификации...";
	Log->Lines->Add("Считываю классификации  всего строк: " + IntToStr(nLastRow-1));
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
				Log->Lines->Add("При считывании найден дубликат - проигнорирован: [" + cl.Name + "]");
            }
		}
	}
	lblStatus->Caption = "Готово";
}

//---------------------------------------------------------------------------

void TForm1::WriteClassifications(Variant &vSheet, std::vector<exlClass> &classes)
{
	lblStatus->Caption = "Приступаю к сортировке классификаций...";
	Log->Lines->Add("Приступаю к сортировке классификаций...");
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
	Log->Lines->Add("Отсортировано!");

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

	lblStatus->Caption = "Произвожу очистку классификаций в файле...";
	Log->Lines->Add("Произвожу очистку классификаций в файле");
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
	lblStatus->Caption = "Приступаю к записи классификаций в файл...";
	Log->Lines->Add("Приступаю к записи в файл " + IntToStr((int)classes.size()));
	ProgressBar1->Max = classes.size();
	ProgressBar1->Min = 0;
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,1).OlePropertySet("Value", WideString("Наименование"));
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,2).OlePropertySet("Value", WideString("Класс"));
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,3).OlePropertySet("Value", WideString("Комментарий"));
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,4).OlePropertySet("Value", WideString("Месяца"));
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

void TForm1::ReadMonth(Variant &vSheet, std::vector<exlMonth> &month)
{
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	int nRowsCount = vSheet.OlePropertyGet("Cells").OlePropertyGet("Rows").OlePropertyGet("Count");
	int nLastRow = vSheet.OlePropertyGet("Cells", nRowsCount, 3).OlePropertyGet("End", xlUp).OlePropertyGet("Row");

	lblStatus->Caption = "Загрузка данных с листа " + strPageName;
	Log->Lines->Add(" * Произвожу загрузку данных с листа " + strPageName + " (строк: " + IntToStr(nLastRow-1) + ")");
	ProgressBar1->Max = nLastRow;
	ProgressBar1->Min = 0;
	int nFound = 0;
	for (int i = 0; i < nLastRow; i++) {
		ProgressBar1->Position = i;
		Application->ProcessMessages();

		exlMonth mon;
		mon.Month = strPageName;
		mon.Name = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,3).OlePropertyGet("Value");
		if (!mon.Name.Trim().IsEmpty()) {
			mon.Price = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OlePropertyGet("Value");
			if (mon.Price != 0) {
				nFound++;
				month.push_back(mon);
			} else {
				Log->Lines->Add("\tЧто то не то: " + mon.Name + " : " + FloatToStr(mon.Price));
			}
		}
	}
	ProgressBar1->Position = 0;
	Log->Lines->Add(" * Найдено записей: " + IntToStr(nFound) + "");
	lblStatus->Caption = "";
}

//---------------------------------------------------------------------------

void TForm1::ReadMonthSum(Variant &vSheet, double &sum)
{
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	int nRowsCount = vSheet.OlePropertyGet("Cells").OlePropertyGet("Rows").OlePropertyGet("Count");
	int nLastRow = vSheet.OlePropertyGet("Cells", nRowsCount, 4).OlePropertyGet("End", xlUp).OlePropertyGet("Row");

	Log->Lines->Add(" * Поиск и суммирования сум по дням " + strPageName + " (строк: " + IntToStr(nLastRow-1) + ")");
	ProgressBar1->Max = nLastRow;
	ProgressBar1->Min = 0;
	int nFound = 0;
	for (int i = 0; i < nLastRow; i++) {
		ProgressBar1->Position = i;
		Application->ProcessMessages();

		exlMonth mon;
		mon.Month = strPageName;
		mon.Name = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,3).OlePropertyGet("Value");
		if (mon.Name.Trim().IsEmpty()) {
			mon.Price = vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4).OlePropertyGet("Value");
			int nWeight = vSheet.
				OlePropertyGet("Cells").
				OlePropertyGet("Item",i+1,4).
				OlePropertyGet("Borders", xlEdgeLeft).
				OlePropertyGet("Weight");
			if (mon.Price != 0 && nWeight == xlMedium) {
				Log->Lines->Add("\tСумма: " + mon.Name + ": " + FloatToStr(mon.Price));
				sum += mon.Price;
			}
		}
	}

	ProgressBar1->Position = 0;
	Log->Lines->Add(" * Найдено: " + IntToStr(nFound) + ", общая сумма: " + FloatToStr(sum));
}

//---------------------------------------------------------------------------

void TForm1::setBorders(Variant &vSheet, int nRow, int nCol) {
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);
}

//---------------------------------------------------------------------------

void TForm1::clearCell(Variant &vSheet, int nRow, int nCol) {
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertySet("Value", WideString(""));
}

//---------------------------------------------------------------------------

void TForm1::setColor(Variant &vSheet, int nRow, int nCol, int nColor) {
	vSheet.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,nCol).OlePropertyGet("Interior").OlePropertySet("Color", nColor);
}

//---------------------------------------------------------------------------

int TForm1::RGBToInt(int r, int g, int b) {
	TColor cl = RGB (r, g, b);
	return cl;
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
	UnicodeString strResult = "=ГИПЕРССЫЛКА(\"#'классификации'!A" + sLine + ":B" + sLine + "\", \"" + class2_ + ": " + name2 + "\")";
	return strResult;
};
//---------------------------------------------------------------------------

void __fastcall TForm1::actCalcClassificationExecute(TObject *Sender)
{
	if(!MakeBackup()) {
		Log->Lines->Add("Не удалось создать резервную копию файла");
		return;
	}

	Variant app = Variant::CreateObject("Excel.Application");
	app.OlePropertySet("Visible", true);
	Variant excel = app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(m_strFileName.c_str()));
	Variant vSheets = excel.OlePropertyGet("Worksheets");
	Variant vSheet = vSheets.OlePropertyGet("Item",m_nPageClassification);
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	if (strPageName.UpperCase() != UnicodeString("классификации").UpperCase()) {
	   app.OleProcedure("Quit");
	   MessageBox (Handle, UnicodeString(L"Не верное имя страницы").c_str(), L"prompt", MB_OK);
	   return;
	};

	std::vector<exlClass> classes;
	ReadClassifications(vSheet, classes);

	std::vector<exlMonth> months;
	Variant vSheetMonth;
	double fSumSum = 0;
	for (unsigned int i = 0; i < m_vMonth.size(); i++) {
		if (cmbMonth->Text == m_vMonth[i].Name) {
			int nMonthPage = m_vMonth[i].Number;
			vSheetMonth = vSheets.OlePropertyGet("Item", nMonthPage);
			ReadMonth(vSheetMonth, months);
			ReadMonthSum(vSheetMonth, fSumSum);
		}
	}

	Log->Lines->Add("Произвожу расчет...");
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
				months[i].LinkToClassification = createHyperLinkToClassification(classes, iC);
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

	Log->Lines->Add("Готово.");

	Log->Lines->Add("Сортирую классификации...");
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
	Log->Lines->Add("Готово.");

	Log->Lines->Add("Очистка старых данных");
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


	Log->Lines->Add("Запись новых данных");
	{
		double nSum = 0;
		int nRow = 2;
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertySet("Value", WideString("Класс"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("Value", WideString("Сумма"));
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

			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));

			setBorders(vSheetMonth, nRow, 11);
			setBorders(vSheetMonth, nRow, 12);
		}
		nRow++;
		ProgressBar1->Position = 0;

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertySet("Value", WideString("Итого:"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("Value", WideString(nSum));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));

		nRow++;

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,11).OlePropertySet("Value", WideString("Сумма сумм по дням:"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("Value", WideString(fSumSum));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,12).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
	}

	{
		double nSum = 0;
		int nRow = 2;
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertySet("Value", WideString("Класс"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OlePropertySet("Value", WideString("Наименование"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("Value", WideString("Цена"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("Value", WideString("Ссылка"));

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertyGet("Font").OlePropertySet("Bold", true);
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertyGet("Font").OlePropertySet("Bold", true);

		vSheetMonth.OlePropertyGet("Columns", WideString("N")).OlePropertySet("ColumnWidth", 20);
		vSheetMonth.OlePropertyGet("Columns", WideString("O")).OlePropertySet("ColumnWidth", 50);
		vSheetMonth.OlePropertyGet("Columns", WideString("P")).OlePropertySet("ColumnWidth", 15);
		vSheetMonth.OlePropertyGet("Columns", WideString("Q")).OlePropertySet("ColumnWidth", 50);

		setBorders(vSheetMonth, nRow, 14);
		setBorders(vSheetMonth, nRow, 15);
		setBorders(vSheetMonth, nRow, 16);
		setBorders(vSheetMonth, nRow, 17);

		setColor(vSheetMonth, nRow, 14, RGBToInt(240, 230, 140));
		setColor(vSheetMonth, nRow, 15, RGBToInt(240, 230, 140));
		setColor(vSheetMonth, nRow, 16, RGBToInt(240, 230, 140));
		setColor(vSheetMonth, nRow, 17, RGBToInt(240, 230, 140));

		ProgressBar1->Max = months.size();
		ProgressBar1->Min = 0;
		for (unsigned int i = 0; i < months.size(); i++) {
			nRow++;
			ProgressBar1->Position = i;
			nSum += months[i].Price;

			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertySet("Value", WideString(months[i].Class));
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,15).OlePropertySet("Value", WideString(months[i].Name));
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("Value", WideString(months[i].Price));
			vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));

			if (months[i].Price < 0) {
				setColor(vSheetMonth, nRow, 16, RGBToInt(240, 230, 140));
			}


			// TODO: check linkt to classification
			// vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("Value", WideString(months[i].LinkToClassification.c_str()));

			setBorders(vSheetMonth, nRow, 14);
			setBorders(vSheetMonth, nRow, 15);
			setBorders(vSheetMonth, nRow, 16);
			setBorders(vSheetMonth, nRow, 17);
		}
		nRow++;
		ProgressBar1->Position = 0;

		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,14).OlePropertySet("Value", WideString("Итого:"));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("Value", WideString(nSum));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,16).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
		vSheetMonth.OlePropertyGet("Cells").OlePropertyGet("Item",nRow,17).OlePropertySet("NumberFormat", WideString(m_strNumberFormat));
	}

	Log->Lines->Add("Сохраняю файл...");
	try {
		app.OlePropertySet("DisplayAlerts",false);
		excel.OleProcedure("SaveAs", WideString(m_strFileName.c_str()));
		Log->Lines->Add("Данные сохранены!");
	} catch (...) {
		Log->Lines->Add("Ошибка: Пожалуйста закройте все открытые копии файла и повторите операцию");
	}
    ProgressBar1->Position = 0;
	app.OleProcedure("Quit");
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
	if (strPageName.UpperCase() != UnicodeString("классификации").UpperCase()) {
	   app.OleProcedure("Quit");
	   MessageBox (Handle, UnicodeString(L"Не верное имя страницы").c_str(), L"prompt", MB_OK);
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
		Log->Lines->Add("Не удалось создать резервную копию файла");
		lblStatus->Caption = "Ошибка: Не удалось создать резервную копию файла";
		return;
	}
	m_strRecomendations = "";

	lblStatus->Caption = "Открываю файл на чтение...";
	Variant app = Variant::CreateObject("Excel.Application");
    app.OlePropertySet("Visible", true);
	Variant excel = app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(m_strFileName.c_str()));
	Variant vSheets = excel.OlePropertyGet("Worksheets");
	Variant vSheet = vSheets.OlePropertyGet("Item",m_nPageClassification);
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	if (strPageName.UpperCase() != UnicodeString("классификации").UpperCase()) {
		app.OleProcedure("Quit");
		MessageBox (Handle, UnicodeString(L"Не верное имя страницы").c_str(), L"prompt", MB_OK);
		lblStatus->Caption = "Ошибка: Не верное имя страницы";
		return;
	};

	std::vector<exlClass> classes;
	ReadClassifications(vSheet, classes);

	std::vector<exlMonth> months;
	for (unsigned int i = 0; i < m_vMonth.size(); i++) {
		int nMonthPage = m_vMonth[i].Number;
		Variant vSheetMonth = vSheets.OlePropertyGet("Item",nMonthPage);
		ReadMonth(vSheetMonth, months);
	}
//	Log->Lines->Add("Обновляю линки." + IntToStr((int)months.size()));

	int nAll = classes.size() * months.size();
	Log->Lines->Add("Всего записей в месяцах: " + IntToStr((int)months.size()));
	std::vector<exlClass> newclasses;
	int nRemovedClasses = 0;
	ProgressBar1->Max = nAll;
	ProgressBar1->Min = 0;
	ProgressBar1->Position = 0;

	lblStatus->Caption = "Поиск наименований которые отсутвуют в месяцах...";
	Log->Lines->Add("Произвожу поиск наименований которые отсутвуют в месяцах...");
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
			Log->Lines->Add("\tНаименование '" + cl.Name + "' - нигде не встречается и будет удалено из классификаций");
		}
	}
	Log->Lines->Add(" ** ");

	nAll = newclasses.size() * months.size();
	ProgressBar1->Max = nAll;
	ProgressBar1->Min = 0;
	ProgressBar1->Position = 0;
	int nAddClasses = 0;
	lblStatus->Caption = "Поиск наименований которые отсутвуют в списке классификаций...";
	Log->Lines->Add("Произвожу поиск наименований которые отсутвуют в списке классификаций...");
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
			Log->Lines->Add("\tНаименование '" + months[iM].Name + "' - будет добавлено в классификации");
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
	Log->Lines->Add("Будет удалено классификаций: " + IntToStr(nRemovedClasses));
    Log->Lines->Add("Будет добавлено классификаций: " + IntToStr(nAddClasses));
	Log->Lines->Add("Всего классификаций: " + IntToStr((int)(newclasses.size())));


	WriteClassifications(vSheet, newclasses);

	lblStatus->Caption = "Сохраняю файл...";
	Log->Lines->Add("Сохраняю файл...");
	try {
		app.OlePropertySet("DisplayAlerts",false);
		excel.OleProcedure("SaveAs", WideString(m_strFileName.c_str()));
		Log->Lines->Add("Данные сохранены!");
	} catch (...) {
		lblStatus->Caption = "Ошибка: Пожалуйста закройте все открытые копии файла и повторите операцию";
		Log->Lines->Add("Ошибка: Пожалуйста закройте все открытые копии файла и повторите операцию");
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
		Log->Lines->Add("Ошибка: Не удалось создать резервную копию файла");
		return;
	}

	Variant app = Variant::CreateObject("Excel.Application");
	app.OlePropertySet("Visible", true);
	Variant excel = app.OlePropertyGet("Workbooks").OleFunction("Open", WideString(m_strFileName.c_str()));
	Variant vSheets = excel.OlePropertyGet("Worksheets");

	Variant vSheet = vSheets.OlePropertyGet("Item",m_nPageClassification);
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	if (strPageName.UpperCase() != UnicodeString("классификации").UpperCase()) {
	   app.OleProcedure("Quit");
	   Log->Lines->Add("Ошибка: Не верное имя страницы");
	   MessageBox (Handle, UnicodeString(L"Не верное имя страницы").c_str(), L"prompt", MB_OK);
	   return;
	};
	std::vector<exlClass> classes;
	ReadClassifications(vSheet, classes);
	WriteClassifications(vSheet, classes);


	Log->Lines->Add("Сохраняю файл...");
	try {
		app.OlePropertySet("DisplayAlerts",false);
		excel.OleProcedure("SaveAs", WideString(m_strFileName.c_str()));
		Log->Lines->Add("Классификации отсортированы!");
	} catch (...) {
		Log->Lines->Add("Ошибка: Пожалуйста закройте все открытые копии файла и повторите операцию");
	}
	app.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormCreate(TObject *Sender)
{
	lblStatus->Caption = "";
}
//---------------------------------------------------------------------------

