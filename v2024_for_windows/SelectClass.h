//---------------------------------------------------------------------------

#ifndef SelectClassH
#define SelectClassH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
//---------------------------------------------------------------------------
class TFormSelectClass : public TForm
{
__published:	// IDE-managed Components
	TComboBox *ComboBox1;
	TButton *Button1;
	TButton *Button2;
	TLabel *Label1;
	void __fastcall Button2Click(TObject *Sender);
	void __fastcall Button1Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TFormSelectClass(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormSelectClass *FormSelectClass;
//---------------------------------------------------------------------------
#endif
