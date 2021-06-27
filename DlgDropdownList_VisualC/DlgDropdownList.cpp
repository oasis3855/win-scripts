
// DlgDropdownList.cpp : アプリケーションのクラス動作を定義します。
//

#include "pch.h"
#include "framework.h"
#include "DlgDropdownList.h"
#include "DlgDropdownListDlg.h"
#include "CDlgMessagebox.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CDlgDropdownListApp

BEGIN_MESSAGE_MAP(CDlgDropdownListApp, CWinApp)
    ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CDlgDropdownListApp の構築

CDlgDropdownListApp::CDlgDropdownListApp()
{
    // TODO: この位置に構築用コードを追加してください。
    // ここに InitInstance 中の重要な初期化処理をすべて記述してください。

    // プログラムのexitコード（戻り値）
    this->numSelect = 0;
}


// 唯一の CDlgDropdownListApp オブジェクト

CDlgDropdownListApp theApp;


// CDlgDropdownListApp の初期化

BOOL CDlgDropdownListApp::InitInstance()
{
    CWinApp::InitInstance();

    // 標準初期化
    // これらの機能を使わずに最終的な実行可能ファイルの
    // サイズを縮小したい場合は、以下から不要な初期化
    // ルーチンを削除してください。
    // 設定が格納されているレジストリ キーを変更します。
    // TODO: 会社名または組織名などの適切な文字列に
    // この文字列を変更してください。
    // SetRegistryKey(_T("アプリケーション ウィザードで生成されたローカル アプリケーション"));

    // 引数の切り分けと配列への格納
    this->arrParams.RemoveAll();
    this->ParseProgramArguments(&this->arrParams);

    // 引数の数により、表示するダイアログを切り替える
    if (this->arrParams.GetSize() < 2)
    {
        // 引数が1個以下の場合は、YES/NO選択ダイアログを表示
        // （引数が無い場合の、ヘルプダイアログもここで処理する）
        CDlgMessagebox dlg;
        m_pMainWnd = &dlg;

        dlg.arrParamsPtr = &this->arrParams;

        INT_PTR nResponse = dlg.DoModal();
        if (nResponse == IDOK)
        {
            // TODO: ダイアログが <OK> で消された時のコードを
            //  記述してください。
            this->numSelect = 1;
        }
        else
        {
            // TODO: ダイアログが <キャンセル> で消された時のコードを
            //  記述してください。
            this->numSelect = 0;
        }
    }
    else
    {
        // 引数が2個以上の場合は、リストボックス選択ダイアログを表示
        CDlgDropdownListDlg dlg;
        m_pMainWnd = &dlg;

        dlg.arrParamsPtr = &this->arrParams;

        INT_PTR nResponse = dlg.DoModal();
        if (nResponse == IDOK)
        {
            // TODO: ダイアログが <OK> で消された時のコードを
            //  記述してください。
            this->numSelect = dlg.numSelected;
        }
        else if (nResponse == IDCANCEL)
        {
            // TODO: ダイアログが <キャンセル> で消された時のコードを
            //  記述してください。
            this->numSelect = 0;
        }
        else if (nResponse == -1)
        {
            this->numSelect = 0;
            //TRACE(traceAppMsg, 0, "警告: ダイアログの作成に失敗しました。アプリケーションは予期せずに終了します。\n");
        }
    }


#if !defined(_AFXDLL) && !defined(_AFX_NO_MFC_CONTROLS_IN_DIALOGS)
    ControlBarCleanUp();
#endif

    // ダイアログは閉じられました。アプリケーションのメッセージ ポンプを開始しないで
    //  アプリケーションを終了するために FALSE を返してください。
    return FALSE;
}



// 引数の切り分けと配列への格納
int CDlgDropdownListApp::ParseProgramArguments(CStringArray* arrParams)
{
    // TODO: ここに実装コードを追加します.
    // 配列を初期化する
    if (!arrParams->IsEmpty()) arrParams->RemoveAll();

    CString cmdParam(m_lpCmdLine);
    cmdParam.Trim();

    CString param;
    int curPos = 0;

    do
    {
        if (cmdParam.GetLength() >= curPos && cmdParam.GetAt(curPos) == ' ')
        {
            // "..." を切り出した次の引数解析は、curPosが空白文字を指すので、文字を一つ先に進める
            ++curPos;
            continue;
        }
        else if (cmdParam.GetLength() >= curPos && cmdParam.GetAt(curPos) == '\"')
        {
            // "で括われた引数の場合
            ++curPos;
            param = cmdParam.Tokenize(_T("\""), curPos);
        }
        else
        {
            // 引数を半角スペースで分解
            param = cmdParam.Tokenize(_T(" "), curPos);
        }
        // 取り出された文字列が空の場合（最後の引数がこれに当たる）は引数配列に追加しない
        if (param.GetLength() > 0)
            arrParams->Add(param.Trim());
    } while (param != "");

    for (int i = 0; i < arrParams->GetSize(); i++)
    {
        param = arrParams->GetAt(i);
    }

    return 0;
}

// プログラムのexitコードをカスタマイズする（デフォルト関数をオーバーライド）

int CDlgDropdownListApp::ExitInstance()
{
    // TODO: ここに特定なコードを追加するか、もしくは基底クラスを呼び出してください。

//	return CWinApp::ExitInstance();

    CWinApp::ExitInstance();
    return this->numSelect;
}
