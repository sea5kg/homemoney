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
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Button1Click(TObject *Sender)
{
	Edit1->Text = "";
	m_strFileName = "";
	if (OpenDialog1->Execute()) {
		m_bBackup = false;
		Button2->Enabled = false;
		Button3->Enabled = false;
		Button5->Enabled = false;
		Button6->Enabled = false;
		Button7->Enabled = false;

		UnicodeString strFileName = OpenDialog1->FileName;
		if (! FileExists (strFileName)) {
			MessageBox (Handle, UnicodeString(L"Файл '" + strFileName + L"' не существует").c_str(), L"prompt", MB_OK);
			Edit1->Text = "";
			return;
		}

		Edit1->Text = strFileName;
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
			Edit1->Text = "";
			return;
		}

		Log->Lines->Add("Файл загружен производиться анализ");
		Variant vSheets = excel.OlePropertyGet("Worksheets");

		m_nPageClassification = 0;
		m_vMonth.clear();
		ComboBox1->Items->Clear();
		int nSheets = vSheets.OlePropertyGet("Count");
		Log->Lines->Add("Всего листов: " + IntToStr(nSheets));
		for (int i = 0; i < nSheets; i++) {
			Variant vSheet = vSheets.OlePropertyGet("Item",i+1);
			UnicodeString str = vSheet.OlePropertyGet("Name");
			if (str.UpperCase() == UnicodeString("классификации").UpperCase()) {
				m_nPageClassification = i+1;
			};
			if (str.UpperCase().Pos("МЕСЯЦ ") > 0) {
				ComboBox1->Items->Add(str);
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
			Edit1->Text = "";
			return;
		}

		if (m_vMonth.size() == 0) {
			MessageBox (Handle, UnicodeString(L"Не найден ни один лист с 'месяц xx'").c_str(), L"prompt", MB_OK);
			app.OleProcedure("Quit");
			Edit1->Text = "";
			return;
        }
		app.OleProcedure("Quit");
//		Button2->Enabled = true;
//		Button3->Enabled = true;
		Button5->Enabled = true;
		Button6->Enabled = true;
		Button7->Enabled = true;
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
		ExtractFileDir(m_strFileName) + L"\\backups\\" +
		newfilename + " (" + TDateTime::CurrentDateTime().FormatString("yyyy-mm-dd hh:nn") + ")" +
		ExtractFileExt(m_strFileName);
	Log->Lines->Add("Сохраняю резервную копию файла в " + newfilename);
	return CopyFile(m_strFileName.c_str(),newfilename.c_str(),false);
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
              cl.Class = "неизвестно";
            }
			classes.push_back(cl);
		}
	}
}

//---------------------------------------------------------------------------

void TForm1::WriteClassifications(Variant &vSheet, std::vector<exlClass> &classes)
{
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
}

//---------------------------------------------------------------------------

void TForm1::ReadMonth(Variant &vSheet, std::vector<exlMonth> &month)
{
	UnicodeString strPageName = vSheet.OlePropertyGet("Name");
	int nRowsCount = vSheet.OlePropertyGet("Cells").OlePropertyGet("Rows").OlePropertyGet("Count");
	int nLastRow = vSheet.OlePropertyGet("Cells", nRowsCount, 3).OlePropertyGet("End", xlUp).OlePropertyGet("Row");

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
	Log->Lines->Add(" * Найдено записей: " + IntToStr(nFound) + "");

}

//---------------------------------------------------------------------------

void __fastcall TForm1::Button6Click(TObject *Sender)
{
	if(!MakeBackup()) {
		Log->Lines->Add("Ошибка: Не удалось создать резервную копию файла");
		return;
	}

	Variant app = Variant::CreateObject("Excel.Application");
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
void __fastcall TForm1::Button5Click(TObject *Sender)
{
	if(!MakeBackup()) {
		Log->Lines->Add("Не удалось создать резервную копию файла");
		return;
	}

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

	std::vector<exlMonth> months;
	for (unsigned int i = 0; i < m_vMonth.size(); i++) {
		int nMonthPage = m_vMonth[i].Number;
		Variant vSheetMonth = vSheets.OlePropertyGet("Item",nMonthPage);
		ReadMonth(vSheetMonth, months);
	}

	int nAll = classes.size() * months.size();
	Log->Lines->Add("Всего записей в месяцах: " + IntToStr((int)months.size()));
	std::vector<exlClass> newclasses;
	int nRemovedClasses = 0;
	ProgressBar1->Max = nAll;
	ProgressBar1->Min = 0;
	ProgressBar1->Position = 0;
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
			cl.Class = "неизвестно";
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
void __fastcall TForm1::Button7Click(TObject *Sender)
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

	FormClassEditor->ShowModal();

}
//---------------------------------------------------------------------------
