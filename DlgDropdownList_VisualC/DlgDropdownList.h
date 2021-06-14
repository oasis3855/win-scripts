
// DlgDropdownList.h : PROJECT_NAME アプリケーションのメイン ヘッダー ファイルです
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH に対してこのファイルをインクルードする前に 'pch.h' をインクルードしてください"
#endif

#include "resource.h"		// メイン シンボル


// CDlgDropdownListApp:
// このクラスの実装については、DlgDropdownList.cpp を参照してください
//

class CDlgDropdownListApp : public CWinApp
{
public:
	CDlgDropdownListApp();

// オーバーライド
public:
	virtual BOOL InitInstance();

// 実装

	DECLARE_MESSAGE_MAP()
	int ParseProgramArguments(CStringArray* arrParams);
	virtual int ExitInstance();
	int numSelect;
	CStringArray arrParams;
};

extern CDlgDropdownListApp theApp;
