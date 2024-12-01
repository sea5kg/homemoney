//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "classeditor.h"
#include "helpers.h"
#include "selectclass.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFormClassEditor *FormClassEditor;
//---------------------------------------------------------------------------
__fastcall TFormClassEditor::TFormClassEditor(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormClassEditor::N2Click(TObject *Sender)
{
	if (TreeView1->Selected->Data != NULL)
		TreeView1->Selected->Delete();
}
//---------------------------------------------------------------------------

void __fastcall TFormClassEditor::N1Click(TObject *Sender)
{
	if (TreeView1->Selected->Data != NULL)
	{
		exlClass *exl = (exlClass *)TreeView1->Selected->Data;
		FormSelectClass->Label1->Caption = exl->Name;
		FormSelectClass->ComboBox1->Text = exl->Class;
		if (FormSelectClass->ShowModal() == mrOk) {
			exl->Class = FormSelectClass->ComboBox1->Text;
//			TreeView1->Selected->Parent =
		}
	}
}
//---------------------------------------------------------------------------

