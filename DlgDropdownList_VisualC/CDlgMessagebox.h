#pragma once


// CDlgMessagebox ダイアログ

class CDlgMessagebox : public CDialog
{
	DECLARE_DYNAMIC(CDlgMessagebox)

public:
	CDlgMessagebox(CWnd* pParent = nullptr);   // 標準コンストラクター
	virtual ~CDlgMessagebox();

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DLGYESNO };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV サポート

	DECLARE_MESSAGE_MAP()
public:
	CStringArray *arrParamsPtr;
	virtual BOOL OnInitDialog();
};
