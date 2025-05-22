//---------------------------------------------------------------------------

#ifndef classeditorH
#define classeditorH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ComCtrls.hpp>
#include <Vcl.ImgList.hpp>
#include <Vcl.Menus.hpp>
#include <Vcl.ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TFormClassEditor : public TForm
{
__published:	// IDE-managed Components
	TImageList *ImageList1;
	TPopupMenu *PopupMenu1;
	TMenuItem *N1;
	TMenuItem *N2;
	TPanel *Panel1;
	TTreeView *TreeView2;
	TLabel *Label2;
	TPanel *Panel2;
	TTreeView *TreeView1;
	TLabel *Label1;
	TPanel *Panel3;
	TSplitter *Splitter1;
	TSplitter *Splitter2;
	void __fastcall N2Click(TObject *Sender);
	void __fastcall N1Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TFormClassEditor(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormClassEditor *FormClassEditor;
//---------------------------------------------------------------------------
#endif
