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
#include <System.Actions.hpp>
#include <Vcl.ActnList.hpp>
#include <vector>
#include      <ComObj.hpp>
#include      <utilcls.h>


const int xlContinuous = 1;
const int xlUp = 3;
const int xlDash = -4115;
const int xlDashDot = 4;
const int xlDashDotDot = 5;
const int xlDot = -4118;
const int xlDouble = -4119;
const int xlLIneStyleNone = -4142;
const int xlSlantDashDot = 13;
const int xlHairline = 1;
const int xlThin = 2;
const int xlMedium = -4138;
const int xlThick = 4;
const int xlDiagonalDown = 5;
const int xlDiagonalUp = 6;
const int xlEdgeBottom = 9;
const int xlEdgeLeft = 7;
const int xlEdgeRight = 10;
const int xlEdgeTop = 8;
const int xlInsideHorizontal = 12;
const int xlInsideVertical = 11;

//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
	TOpenDialog *OpenDialog1;
	TEdit *edtFile;
	TButton *Button1;
	TMemo *Log;
	TProgressBar *ProgressBar1;
	TButton *btnViewClassifications;
	TActionList *ActionList1;
	TAction *actCalcClassification;
	TAction *actViewClassifications;
	TAction *actRedesignClassifications;
	TAction *actSortClassifications;
	TGroupBox *GroupBox1;
	TComboBox *cmbMonth;
	TButton *btnCalcClassifications;
	TGroupBox *GroupBox2;
	TButton *btnRedesignClassifications;
	TButton *btnSortClassifications;
	TLabel *lblStatus;
	void __fastcall Button1Click(TObject *Sender);
	void __fastcall actCalcClassificationExecute(TObject *Sender);
	void __fastcall actCalcClassificationUpdate(TObject *Sender);
	void __fastcall actViewClassificationsExecute(TObject *Sender);
	void __fastcall actViewClassificationsUpdate(TObject *Sender);
	void __fastcall actRedesignClassificationsUpdate(TObject *Sender);
	void __fastcall actRedesignClassificationsExecute(TObject *Sender);
	void __fastcall actSortClassificationsUpdate(TObject *Sender);
	void __fastcall actSortClassificationsExecute(TObject *Sender);
	void __fastcall FormCreate(TObject *Sender);
private:	// User declarations
	UnicodeString m_sUnknownClass;
	UnicodeString m_strDecDelim;
	UnicodeString m_strNumberFormat;
	int m_nPageClassification;
	std::vector<exlSheet> m_vMonth;
	bool m_bBackup;
	UnicodeString m_strFileName;
	bool MakeBackup();
	void ReadClassifications(Variant &vSheet, std::vector<exlClass> &classes);
	void WriteClassifications(Variant &vSheet, std::vector<exlClass> &classes);
	void ReadMonth(Variant &vSheet, std::vector<exlMonth> &month);
	void ReadMonthSum(Variant &vSheet, double &sum);
	void setBorders(Variant &vSheet, int nRow, int nCol);
	void setColor(Variant &vSheet, int nRow, int nCol, int nColor);
	void clearCell(Variant &vSheet, int nRow, int nCol);
	int RGBToInt(int r, int g, int b);
    UnicodeString createHyperLinkToClassification(std::vector<exlClass> &classes, int nLine);

public:		// User declarations
	__fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
