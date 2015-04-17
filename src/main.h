//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.Dialogs.hpp>
#include "helpers.h"
#include <Vcl.ComCtrls.hpp>
#include <vector>
#include      <ComObj.hpp>
#include      <utilcls.h>


int xlUp = 3;

//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
	TOpenDialog *OpenDialog1;
	TEdit *Edit1;
	TButton *Button1;
	TLabel *Label1;
	TLabel *Label2;
	TButton *Button2;
	TButton *Button3;
	TButton *Button5;
	TButton *Button6;
	TComboBox *ComboBox1;
	TMemo *Log;
	TProgressBar *ProgressBar1;
	TButton *Button7;
	void __fastcall Button1Click(TObject *Sender);
	void __fastcall Button6Click(TObject *Sender);
	void __fastcall Button5Click(TObject *Sender);
	void __fastcall Button7Click(TObject *Sender);
private:	// User declarations
	int m_nPageClassification;
	std::vector<exlSheet> m_vMonth;
	bool m_bBackup;
	UnicodeString m_strFileName;
	bool MakeBackup();
	void ReadClassifications(Variant &vSheet, std::vector<exlClass> &classes);
	void WriteClassifications(Variant &vSheet, std::vector<exlClass> &classes);
	void TForm1::ReadMonth(Variant &vSheet, std::vector<exlMonth> &month);

public:		// User declarations
	__fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
