
// aaaaDlg.h: 헤더 파일
//

#pragma once


// CaaaaDlg 대화 상자
class CaaaaDlg : public CDialogEx
{
// 생성입니다.
public:
	CaaaaDlg(CWnd* pParent = nullptr);	// 표준 생성자입니다.

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_AAAA_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.


// 구현입니다.
protected:
	HICON m_hIcon;

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CListCtrl m_listview;
	CString m_id;
	CString m_subject;
	int m_hours;
	int m_nSelected;
	CComboBox m_comboFind;
	CString m_strFind;
	CMFCEditBrowseCtrl m_browse1;
	CString m_name;
	afx_msg void OnBnClickedMfcbuttondboepn();
	afx_msg void OnBnClickedMfcbuttonFirst();
	afx_msg void OnBnClickedMfcbuttonPrev();
	afx_msg void OnBnClickedMfcbuttonNext();
	afx_msg void OnBnClickedMfcbuttonLast();
	void ImageUpdate();
	afx_msg void OnEnChangeMfceditbrowse1Pic1();
	CStatic m_picImage;
	afx_msg void OnBnClickedMfcbuttonInput();
	afx_msg void OnBnClickedMfcbuttonModify();
	afx_msg void OnBnClickedMfcbuttonDel();
	afx_msg void OnBnClickedMfcbuttonClear();
	afx_msg void OnCbnSelchangeComboFind();
	afx_msg void OnBnClickedMfcbuttonFind();
	afx_msg void OnStnClickedStaticPic1();
	afx_msg void OnBnClickedMfcbutton12();
	afx_msg void OnBnClickedMfcbuttonExcel();
	CListCtrl m_listView2;
	CString strPathName;
	afx_msg void OnBnClickedMfcbuttonOrder();
	afx_msg void OnBnClickedMfcbuttonClear2();
	CString  strSum;
	CString	strCount;
	int TotalPrice;
	int TotalCount;
	afx_msg void OnBnClickedMfcbuttonDel2();
	CComboBox m_comboTable;
	CListBox m_listTable;
	CString strTableNo;
	afx_msg void OnBnClickedListprint();
	afx_msg void OnSelchangeCombo1();
	CComboBox m_comboCnt;
	afx_msg void OnCbnSelchangeCombo2();
	CStatic m_modify;
	afx_msg void OnBnClickedMfcbuttonExcel2();
};
