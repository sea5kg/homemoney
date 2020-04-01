//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "SelectClass.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFormSelectClass *FormSelectClass;
//---------------------------------------------------------------------------
__fastcall TFormSelectClass::TFormSelectClass(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormSelectClass::Button2Click(TObject *Sender)
{
	Close();
}
//---------------------------------------------------------------------------
void __fastcall TFormSelectClass::Button1Click(TObject *Sender)
{
	Close();
}
//---------------------------------------------------------------------------
