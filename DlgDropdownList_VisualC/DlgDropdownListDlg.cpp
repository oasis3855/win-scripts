
// DlgDropdownListDlg.cpp : 実装ファイル
//

#include "pch.h"
#include "framework.h"
#include "DlgDropdownList.h"
#include "DlgDropdownListDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CDlgDropdownListDlg ダイアログ



CDlgDropdownListDlg::CDlgDropdownListDlg(CWnd* pParent /*=nullptr*/)
    : CDialog(IDD_DLGDROPDOWNLIST_DIALOG, pParent)
{
    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

    // 初期値
    numSelected = 0;
    arrParamsPtr = NULL;	// 「初期化していない」警告が出るための処理
}

void CDlgDropdownListDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_LIST_SELECTION, listSelection);
}

BEGIN_MESSAGE_MAP(CDlgDropdownListDlg, CDialog)
    ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(IDOK, &CDlgDropdownListDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CDlgDropdownListDlg メッセージ ハンドラー

BOOL CDlgDropdownListDlg::OnInitDialog()
{
    CDialog::OnInitDialog();

    // このダイアログのアイコンを設定します。アプリケーションのメイン ウィンドウがダイアログでない場合、
    //  Framework は、この設定を自動的に行います。
    SetIcon(m_hIcon, TRUE);			// 大きいアイコンの設定
    SetIcon(m_hIcon, FALSE);		// 小さいアイコンの設定

    // TODO: 初期化をここに追加します。
    if (arrParamsPtr->GetSize() < 2)
    {
        // 引数が1個以下の場合（本来はYES/NOダイアログを表示すべきだが、内部エラーでここの処理に来た可能性がある）
        this->SetDlgItemTextW(IDC_TEXT_MESSAGE, _T("引数は3個以上の必要があります（内部エラー）"));
    }
    else
    {
        // 引数が2個以上の場合、第1引数をメッセージ文字列、第2引数以降をリストボックスにセット

        // メッセージ文字列のセット
        this->SetDlgItemTextW(IDC_TEXT_MESSAGE, arrParamsPtr->GetAt(0));
        // リストボックスのセット
        for (int i = 1; i < arrParamsPtr->GetSize(); i++)
        {
            this->listSelection.AddString(arrParamsPtr->GetAt(i));
        }
        // リストボックスの最初の行を選択状態にする
        this->listSelection.SetCurSel(0);

    }

    return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
}

// ダイアログに最小化ボタンを追加する場合、アイコンを描画するための
//  下のコードが必要です。ドキュメント/ビュー モデルを使う MFC アプリケーションの場合、
//  これは、Framework によって自動的に設定されます。

void CDlgDropdownListDlg::OnPaint()
{
    if (IsIconic())
    {
        CPaintDC dc(this); // 描画のデバイス コンテキスト

        SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

        // クライアントの四角形領域内の中央
        int cxIcon = GetSystemMetrics(SM_CXICON);
        int cyIcon = GetSystemMetrics(SM_CYICON);
        CRect rect;
        GetClientRect(&rect);
        int x = (rect.Width() - cxIcon + 1) / 2;
        int y = (rect.Height() - cyIcon + 1) / 2;

        // アイコンの描画
        dc.DrawIcon(x, y, m_hIcon);
    }
    else
    {
        CDialog::OnPaint();
    }
}

// ユーザーが最小化したウィンドウをドラッグしているときに表示するカーソルを取得するために、
//  システムがこの関数を呼び出します。
HCURSOR CDlgDropdownListDlg::OnQueryDragIcon()
{
    return static_cast<HCURSOR>(m_hIcon);
}


void CDlgDropdownListDlg::OnBnClickedOk()
{
    // TODO: ここにコントロール通知ハンドラー コードを追加します。

    // リストボックスで選択されている項目（1～）をクラス変数に格納
    numSelected = listSelection.GetCurSel() + 1;

    CDialog::OnOK();
}
