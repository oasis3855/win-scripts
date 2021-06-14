// CDlgMessagebox.cpp : 実装ファイル
//

#include "pch.h"
#include "DlgDropdownList.h"
#include "CDlgMessagebox.h"
#include "afxdialogex.h"


// CDlgMessagebox ダイアログ

IMPLEMENT_DYNAMIC(CDlgMessagebox, CDialog)

CDlgMessagebox::CDlgMessagebox(CWnd* pParent /*=nullptr*/)
    : CDialog(IDD_DLGYESNO, pParent)
{
    arrParamsPtr = NULL;	// 「初期化していない」警告が出るための処理
}

CDlgMessagebox::~CDlgMessagebox()
{
}

void CDlgMessagebox::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CDlgMessagebox, CDialog)
END_MESSAGE_MAP()


// CDlgMessagebox メッセージ ハンドラー


BOOL CDlgMessagebox::OnInitDialog()
{
    CDialog::OnInitDialog();

    // TODO: ここに初期化を追加してください

    if (arrParamsPtr->GetCount() < 1)
    {
        // 引数が無い場合は、ヘルプメッセージを表示
        this->SetDlgItemTextW(IDC_TEXT_MESSAGE, _T("YES/NO選択メッセージボックス\n"
            "   引数1 : メッセージ文字列\n"
            "リストボックス選択ダイアログ\n"
            "   引数1 : メッセージ文字列, 引数2～ : リスト項目\n"));
        // 2個あるボタンのうち、1個を消して、残りの表示を「OK」に書き換える
        CButton* btnOk = (CButton*)GetDlgItem(IDOK);
        btnOk->ShowWindow(SW_HIDE);
        CButton* btnCancel = (CButton*)GetDlgItem(IDCANCEL);
        btnCancel->SetWindowTextW(_T("OK"));
        // ダイアログタイトルの設定
        this->SetWindowTextW(_T("このプログラムの使い方"));
    }
    else
    {
        // 引数が1個の場合は、引数1をメッセージエリアにセットする
        this->SetDlgItemTextW(IDC_TEXT_MESSAGE, arrParamsPtr->GetAt(0));
    }


    return TRUE;  // return TRUE unless you set the focus to a control
                  // 例外 : OCX プロパティ ページは必ず FALSE を返します。
}
