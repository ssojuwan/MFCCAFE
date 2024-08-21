
// aaaaDlg.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "aaaa.h"
#include "aaaaDlg.h"
#include "afxdialogex.h"

//db 관련 헤더파일 추가
#include <afxdb.h>

//엑셀 관련 헤더 파일 추가
#include "XLAutomation.h"
#include "XLEzAutomation.h"
#include "database1.h"


database1 m_pSet;


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 응용 프로그램 정보에 사용되는 CAboutDlg 대화 상자입니다.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

// 구현입니다.
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CaaaaDlg 대화 상자



CaaaaDlg::CaaaaDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_AAAA_DIALOG, pParent)
	, m_id(_T(""))
	, m_hours(0)
	, m_nSelected(0)
	, m_strFind(_T(""))
	, m_name(_T(""))
	, strPathName(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CaaaaDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_listview);
	DDX_Text(pDX, IDC_EDIT_ID, m_id);
	DDX_Text(pDX, IDC_EDIT_HOURS, m_hours);
	DDX_CBIndex(pDX, IDC_COMBO_FIND, m_nSelected);
	DDX_Control(pDX, IDC_COMBO_FIND, m_comboFind);
	DDX_Text(pDX, IDC_EDIT_FIND, m_strFind);
	DDX_Control(pDX, IDC_MFCEDITBROWSE1_PIC1, m_browse1);
	DDX_Text(pDX, IDC_EDIT_NAME, m_name);
	DDX_Control(pDX, IDC_STATIC_PIC1, m_picImage);
	DDX_Control(pDX, IDC_LIST2, m_listView2);
	DDX_Text(pDX, IDC_MFCEDITBROWSE1_PIC1, strPathName);

	DDX_Control(pDX, IDC_COMBO1, m_comboTable);
	DDX_Control(pDX, IDC_LIST_TABLE, m_listTable);
	DDX_Control(pDX, IDC_COMBO2, m_comboCnt);
	DDX_Control(pDX, IDC_STATIC_MODIFY, m_modify);
}

BEGIN_MESSAGE_MAP(CaaaaDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_MFCBUTTONdboepn, &CaaaaDlg::OnBnClickedMfcbuttondboepn)
	ON_BN_CLICKED(IDC_MFCBUTTON_FIRST, &CaaaaDlg::OnBnClickedMfcbuttonFirst)
	ON_BN_CLICKED(IDC_MFCBUTTON_PREV, &CaaaaDlg::OnBnClickedMfcbuttonPrev)
	ON_BN_CLICKED(IDC_MFCBUTTON_NEXT, &CaaaaDlg::OnBnClickedMfcbuttonNext)
	ON_BN_CLICKED(IDC_MFCBUTTON_LAST, &CaaaaDlg::OnBnClickedMfcbuttonLast)
	ON_EN_CHANGE(IDC_MFCEDITBROWSE1_PIC1, &CaaaaDlg::OnEnChangeMfceditbrowse1Pic1)
	ON_BN_CLICKED(IDC_MFCBUTTON_INPUT, &CaaaaDlg::OnBnClickedMfcbuttonInput)
	ON_BN_CLICKED(IDC_MFCBUTTON_MODIFY, &CaaaaDlg::OnBnClickedMfcbuttonModify)
	ON_BN_CLICKED(IDC_MFCBUTTON_DEL, &CaaaaDlg::OnBnClickedMfcbuttonDel)
	ON_BN_CLICKED(IDC_MFCBUTTON_CLEAR, &CaaaaDlg::OnBnClickedMfcbuttonClear)
	ON_CBN_SELCHANGE(IDC_COMBO_FIND, &CaaaaDlg::OnCbnSelchangeComboFind)
	ON_BN_CLICKED(IDC_MFCBUTTON_FIND, &CaaaaDlg::OnBnClickedMfcbuttonFind)
	ON_STN_CLICKED(IDC_STATIC_PIC1, &CaaaaDlg::OnStnClickedStaticPic1)
	ON_BN_CLICKED(IDC_MFCBUTTON12, &CaaaaDlg::OnBnClickedMfcbutton12)
	ON_BN_CLICKED(IDC_MFCBUTTON_EXCEL, &CaaaaDlg::OnBnClickedMfcbuttonExcel)
	ON_BN_CLICKED(IDC_MFCBUTTON_ORDER, &CaaaaDlg::OnBnClickedMfcbuttonOrder)
	ON_BN_CLICKED(IDC_MFCBUTTON_DEL2, &CaaaaDlg::OnBnClickedMfcbuttonDel2)
	ON_BN_CLICKED(IDC_MFCBUTTON_CLEAR2, &CaaaaDlg::OnBnClickedMfcbuttonClear2)
	ON_BN_CLICKED(IDC_LISTPRINT, &CaaaaDlg::OnBnClickedListprint)
	ON_CBN_SELCHANGE(IDC_COMBO1, &CaaaaDlg::OnSelchangeCombo1)
	ON_CBN_SELCHANGE(IDC_COMBO2, &CaaaaDlg::OnCbnSelchangeCombo2)
	ON_BN_CLICKED(IDC_MFCBUTTON_EXCEL2, &CaaaaDlg::OnBnClickedMfcbuttonExcel2)
END_MESSAGE_MAP()


// CaaaaDlg 메시지 처리기

BOOL CaaaaDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	MoveWindow(0, 0, 1900, 900, TRUE);

	// 시스템 메뉴에 "정보..." 메뉴 항목을 추가합니다.

	// IDM_ABOUTBOX는 시스템 명령 범위에 있어야 합니다.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.
	// 확장 스타일 지정
	m_listview.SetExtendedStyle(m_listview.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_CHECKBOXES | LVS_EX_TRACKSELECT);
	m_listview.SetExtendedStyle(m_listview.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP);
	m_listView2.SetExtendedStyle(m_listView2.GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_CHECKBOXES | LVS_EX_TRACKSELECT);
	m_listView2.SetExtendedStyle(m_listView2.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP);
	// List View 열 추가
	m_listview.InsertColumn(0, _T("Code"), LVCFMT_LEFT, 80);
	m_listview.InsertColumn(1, _T("Menu"), LVCFMT_LEFT, 130);
	m_listview.InsertColumn(3, _T("Price"), LVCFMT_LEFT, 80);
	m_listview.InsertColumn(4, _T("Images"), LVCFMT_LEFT, 450);

	// List View2 열 추가
	m_listView2.InsertColumn(0, _T("코드"), LVCFMT_LEFT, 90);
	m_listView2.InsertColumn(1, _T("메뉴"), LVCFMT_LEFT, 150);
	m_listView2.InsertColumn(2, _T("가격"), LVCFMT_RIGHT, 70);
	m_listView2.InsertColumn(3, _T("갯수"), LVCFMT_RIGHT, 70);
	m_listView2.InsertColumn(4, _T("계"), LVCFMT_RIGHT, 80);

	// 테이블 NO 콤보 상자 초기화
	for (int ti = 0; ti <= 50; ti++) {
		CString strTableNo;
		strTableNo.Format(_T("%3d"), ti);
		m_comboTable.AddString(strTableNo);
	}

	m_comboTable.SetCurSel(0); // 0번째 기본 설정

	// 주문 갯수 콤보 상자 초기화
	for (int i = 0; i <= 500; i++) {
		CString strOrderCount;
		strOrderCount.Format(_T("%3d"), i);
		m_comboCnt.AddString(strOrderCount);
	}

	m_comboCnt.SetCurSel(1); // 1번째 기본 설정



	m_pSet.Open();

	if (m_pSet.GetRecordCount() > 0)
	{
		m_browse1.SetWindowText(m_pSet.db_m_Image);
	}

	return TRUE;
	// 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

void CaaaaDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 애플리케이션의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CaaaaDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CaaaaDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CaaaaDlg::OnBnClickedMfcbuttondboepn()
{
	// ODBC에 등록된 데이터를 연다
	//m_pSet.Delete();
	m_listview.DeleteAllItems(); // 리스트 컨트롤 화면 지우기

	CDatabase db;

	// 데이터베이스 객체 선언 및 연결
	db.OpenEx(_T("DSN=cafemenu"), 0);  // ODBC 교사별 시간표, 0: 기본값(읽기/쓰기 모두 지원), DSN: Data Source Name
	//db.OpenEx(_T("ODBC;DSN=db1;UID=guest;PWD=1234"));  // ID, 비번 등  

	// 레코드셋 객체 생성, 데이터 소스에서 선택한 레코드 집합을 나타냅니다.
	CRecordset rs(&db);

	// table1 테이블의 모든 내용을 선택
	rs.Open(CRecordset::dynaset, _T("SELECT * FROM table1"));
	// 열기 형식 지정자
	// 1. dynaset: 양방향 스크롤 레코드 셋-레코드 셋의 순서는 레코드 셋이 열릴 때 결정
	//             조회한 데이터베이스의 내용을 다른 사용자가 변경하더라도 변경된 내용을 조회된 
	//             데이터에 동기화시켜주는 방식
	// 2. snapshot: 양방향 스크롤의 정적 레코드셋-다른 사용자가 변경한 내용이 반되지 않음
	// 3. dynamic: 양방향 스크롤의 레코드 셋- 다른 사용자가 변경한 내용이 레코드 셋에 반영
	// 4. forwardOnly: 전방향 스크롤의 일기 전용 레코드 셋

	CString str;
	int i = 0;
	int price = 0; // 수업 시수

	while (!rs.IsEOF())
	{
		rs.GetFieldValue(short(0), str);
		m_listview.InsertItem(i, str);

		rs.GetFieldValue(short(1), str);
		m_listview.SetItemText(i, 1, str);

		str.Format(_T("%d"), price); // 수업 시수 정수 숫자 변환 %d 사용
		rs.GetFieldValue(short(2), str);
		m_listview.SetItemText(i, 2, str);

		rs.GetFieldValue(short(3), str);  // 이미지 불러오기
		m_listview.SetItemText(i, 3, str);

		//m_pSet.GetFieldValue(short(4), strPathName);  // 이미지 저장
		//m_listview.SetItemText(i, 4, m_pSet.db_m_ImageFile);

		rs.MoveNext();  // 데이터베이스의 다음 레코드로 이동

		// MoveNext(): 데이터베이스의 다음 레코드로 이동
		// MoveFirst(): 데이터베이스의 처음 레코드로 이동
		// MoveLast(): 데이터베이스의 맨 마지막 레코드로 이동
		// MovePrev(): 데이터베이스의 이전 레코드로 이동

		// counter 증가
		i++;
	}
	int cnt = m_listview.GetItemCount();
	CString sText;
	sText.Format(TEXT("%d"), cnt);
	SetDlgItemText(IDC_DB_COUNT, sText);
	rs.Close();
	db.Close();
}

void CaaaaDlg::OnBnClickedMfcbuttonFirst()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	m_pSet.MoveFirst();
	m_id = m_pSet.db_m_Code;
	m_name = m_pSet.db_m_Menu;
	m_hours = m_pSet.db_m_Price;
	m_browse1.SetWindowText(m_pSet.db_m_Image);
	ImageUpdate();
	UpdateData(false); // 변수 => 컨트롤에 넣어 갱신,  UpdateData(TRUE): 컨트롤 값 => 변수에 기억후 갱신
}


void CaaaaDlg::OnBnClickedMfcbuttonPrev()
{
	if (m_pSet.IsBOF()) {
		AfxMessageBox(_T("첫 번째 데이터입니다"));
		m_pSet.MoveNext();
	}
	else
	{
		m_pSet.MovePrev();
		m_id = m_pSet.db_m_Code;
		m_name = m_pSet.db_m_Menu;
		m_hours = m_pSet.db_m_Price;
		m_browse1.SetWindowText(m_pSet.db_m_Image);
		ImageUpdate();
		UpdateData(false); // 변수 => 컨트롤에 넣어 갱신,  UpdateData(TRUE): 컨트롤 값 => 변수에 기억후 갱신
	}
}

void CaaaaDlg::OnBnClickedMfcbuttonNext()
{
	if (m_pSet.IsEOF()) {
		AfxMessageBox(_T("마지막 데이터입니다"));
		m_pSet.MovePrev();  // 마지막 레코드이면 이전 레코드로 가시오
	}
	else
	{
		m_pSet.MoveNext();
		m_id = m_pSet.db_m_Code;
		m_name = m_pSet.db_m_Menu;
		m_hours = m_pSet.db_m_Price;
		m_browse1.SetWindowText(m_pSet.db_m_Image);
		ImageUpdate();
		UpdateData(false); // 변수 => 컨트롤에 넣어 갱신,  UpdateData(TRUE): 컨트롤 값 => 변수에 기억후 갱신
	}
}

void CaaaaDlg::OnBnClickedMfcbuttonLast()
{
	m_pSet.MoveLast();
	m_id = m_pSet.db_m_Code;
	m_name = m_pSet.db_m_Menu;
	m_hours = m_pSet.db_m_Price;
	m_browse1.SetWindowText(m_pSet.db_m_Image);
	ImageUpdate();
	UpdateData(false); // 변수 => 컨트롤에 넣어 갱신,  UpdateData(TRUE): 컨트롤 값 => 변수에 기억후 갱신
}

void CaaaaDlg::ImageUpdate()
{
	m_pSet.db_m_Image = strPathName; // 이미지 갱신

	HBITMAP hBitmap = (HBITMAP)LoadImage(::AfxGetApp()->m_hInstance,
		m_pSet.db_m_Image, IMAGE_BITMAP, 300, 300, LR_LOADFROMFILE);

	// 이미지를 픽쳐컨트롤 박스에 출력
	m_picImage.SetBitmap(hBitmap);

	if (hBitmap != NULL)
		::DeleteObject(hBitmap); // 비트맵 핸들을 삭제

	//UpdateData(FALSE); // 변수 => 컨트롤에 넣어 갱신(화면 출력), UpdateData(TRUE): 컨트롤 값 => 변수에 기억후 갱신
}

// 프로젝트-클래스마법사-ㅡCmfc065odbc4Dlg-명령-개체ID(IDC_MFCEDITBROWSE_PIC)-처리 EN_CHANGE 멤버함수 추가
void CaaaaDlg::OnEnChangeMfceditbrowse1Pic1()
{
	m_browse1.GetWindowText(m_pSet.db_m_Image); // DB 이미지가 출력
	ImageUpdate();
}



void CaaaaDlg::OnBnClickedMfcbuttonInput()
{
	UpdateData(TRUE);  // 컨트롤에 있는 값 => 변수에 넣어 갱신

	m_pSet.AddNew();

	CString str;
	int i = 0;

	m_pSet.db_m_Code = m_id;    // 메뉴코드
	m_pSet.db_m_Menu = m_name;    // 음식 메뉴
	m_pSet.db_m_Price = m_hours;  // 가격

	//-------- 리스트 컨트롤에도 추가---------------------------------
	m_pSet.GetFieldValue(short(0), m_id);
	m_listview.InsertItem(i, m_pSet.db_m_Code);

	m_pSet.GetFieldValue(short(1), m_name);
	m_listview.SetItemText(i, 1, m_pSet.db_m_Menu);

	str.Format(_T("%d"), m_hours); // 가격 정수 숫자 변환 %d 사용
	m_pSet.GetFieldValue(short(2), str);
	m_listview.SetItemText(i, 2, str);

	m_pSet.GetFieldValue(short(3), m_pSet.db_m_Image);  // 이미지 가져오기
	m_listview.SetItemText(i, 3, m_pSet.db_m_Image);

	//counter 증가
	i++;
	//----------------------------------------------------------------

	m_browse1.SetWindowText(m_pSet.db_m_Image);
	ImageUpdate();

	m_pSet.Update();
	m_pSet.Requery();  // 레코드 집합의 쿼리를 다시 실행하여 선택한 레코드를 새로 고칩니다.

	UpdateData(FALSE);

	AfxMessageBox(_T("데이터가 추가되었습니다."));

	OnBnClickedMfcbuttondboepn();  // 입력한 후 바로 리스트 컨트롤에 보여주기 위해서
}


void CaaaaDlg::OnBnClickedMfcbuttonModify()
{
	if (m_pSet.IsEOF())
	{
		m_pSet.MoveLast();
	}
	else if (m_pSet.IsBOF())
	{
		m_pSet.MoveFirst();
	}

	UpdateData(TRUE); // 컨트롤에 있는 값 => 변수에 넣어 갱신

	m_pSet.Edit(); // 데이터 수정 함수

	m_pSet.db_m_Code = m_id;
	m_pSet.db_m_Menu = m_name;
	m_pSet.db_m_Price = m_hours;

	m_browse1.SetWindowText(m_pSet.db_m_Image);
	ImageUpdate();

	m_pSet.Update();

	m_pSet.Requery(); // 레코드 집합의 쿼리를 다시 실행하여 선택한 레코드를 새로 고칩니다.

	UpdateData(FALSE);

	AfxMessageBox(_T("데이터가 수정 되었습니다."));
	OnBnClickedMfcbuttondboepn();  // 입력한 후 바로 리스트 컨트롤에 보여주기 위해서
}



void CaaaaDlg::OnBnClickedMfcbuttonDel()
{
	if (MessageBox(_T("현재 레코드를 삭제하시겠습니까?"), _T("레코드 삭제창"), MB_ICONQUESTION | MB_YESNO) == IDNO)
		return;

	// 현재 레코드를 삭제
	m_pSet.Delete();
	AfxMessageBox(_T("레코드가 삭제 되었습니다."));

	// 삭제 후 명시적으로 다음 레코드로 이동
	m_pSet.MoveNext();

	// 나머지 코드
	m_id = m_pSet.db_m_Code;
	m_name = m_pSet.db_m_Menu;
	m_hours = 0;
	m_browse1.SetWindowText(m_pSet.db_m_Image);
	ImageUpdate();
	UpdateData(FALSE);
	OnBnClickedMfcbuttondboepn();  // 입력한 후 바로 리스트 컨트롤에 보여주기 위해서
}




void CaaaaDlg::OnBnClickedMfcbuttonClear()
{
	m_id.Empty();
	m_name.Empty();
	m_hours = 0;
	UpdateData(false); // 변수=> 컨트롤에 출력
}

void CaaaaDlg::OnCbnSelchangeComboFind()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	m_nSelected = ((CComboBox*)GetDlgItem(IDC_COMBO_FIND))->GetCurSel();
}


void CaaaaDlg::OnBnClickedMfcbuttonFind()
{
	TCHAR m_strFind[500];
	memset(m_strFind, 0x00, sizeof(m_strFind));

	::GetDlgItemText(this->m_hWnd, IDC_EDIT_FIND, m_strFind, sizeof(m_strFind));

	switch (m_nSelected)
	{
	case 0:
		m_pSet.m_strFilter.Format(_T("MENU='%s'"), m_strFind);
		break;

	case 1:
		m_pSet.m_strFilter.Format(_T("CODE='%s'"), m_strFind);
		break;
	}

	m_pSet.Requery();

	if (!m_pSet.IsBOF() && !m_pSet.IsEOF())
	{
		m_id = m_pSet.db_m_Code;
		m_name = m_pSet.db_m_Menu;
		m_hours = m_pSet.db_m_Price;
		m_browse1.SetWindowText(m_pSet.db_m_Image);
		ImageUpdate();
		UpdateData(FALSE);
	}
	else
	{
		// 검색 결과가 없으면 초기화
		OnBnClickedMfcbuttonClear();
	}

	// 필터를 비워줌
	m_pSet.m_strFilter.Empty();
	m_pSet.Requery();
}








void CaaaaDlg::OnStnClickedStaticPic1()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}


void CaaaaDlg::OnBnClickedMfcbutton12()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	PostQuitMessage(0);
}


void CaaaaDlg::OnBnClickedMfcbuttonExcel()
{
	int m_iMax;
	int nColumncount = m_listview.GetItemCount();
	int columnNum = 0;
	CString m_strFileName;

	m_iMax = nColumncount;

	CXLEzAutomation XL(FALSE); // 엑셀 API 함수를 사용하기 위한 클래스 변수 선언

	m_strFileName = "카페 메뉴 리스트";

	XL.SetCellValue(++columnNum, 1, _T("Code"));  // SetCellValue: 셀의 내용 설정
	XL.SetCellValue(++columnNum, 1, _T("Menu"));
	XL.SetCellValue(++columnNum, 1, _T("Price"));
	XL.SetCellValue(++columnNum, 1, _T("Images"));

	for (int i = 1; i <= m_iMax; i++)
	{
		XL.SetCellValue(1, i + 1, m_listview.GetItemText(i - 1, 0)); // SetCellValue: 셀의 내용 설정, 0=> 하위 항목 모두 검색

		for (int j = 1; j <= columnNum; j++)
			XL.SetCellValue(j + 1, i + 1, m_listview.GetItemText(i - 1, j));
	}

	CFileDialog dlg(false, _T("xlsx"), m_strFileName + _T(".xlsx"),
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_NOCHANGEDIR,
		_T("xlsx 파일 (*.xlsx)|*.xlsx|)"), NULL);
	// OFN_HIDEREADONLY: 읽기 전용 확인란을 숨깁니다.
	// OFN_OVERWRITEPROMPT:  선택한 파일이 이미 있는 경우 다른 이름으로 저장 대화 상자에서 메시지 상자를 생성합니다.
	// OFN_NOCHANGEDIR: 사용자가 파일을 검색하는 동안 디렉터리를 변경한 경우 현재 디렉터리를 원래 값으로 복원합니다.

	if (dlg.DoModal() == IDOK)
		XL.SaveFileAs(dlg.GetPathName());  // SaveFileAs: 엑셀 파일 저장

	XL.ReleaseExcel(); // 엑셀 파일 종료
}

void CaaaaDlg::OnBnClickedMfcbuttonOrder()
{
	int cnt = m_listview.GetItemCount();

	// 음식 주문 처리
	int OrderCount = 0;
	int OrderPrice = 0;
	CString str;


	for (int i = 0; i < cnt; i++)
	{
		if (m_listview.GetItemState(i, LVIS_SELECTED) != 0 || (m_listview.GetCheck(i)))
		{
			m_pSet.db_m_Code = m_listview.GetItemText(i, 0);   // 코드값
			m_listView2.InsertItem(0, m_pSet.db_m_Code);

			m_pSet.db_m_Menu = m_listview.GetItemText(i, 1);  // 메뉴값
			m_listView2.SetItemText(0, 1, m_pSet.db_m_Menu);

			m_pSet.db_m_Price = _ttoi(m_listview.GetItemText(i, 2));  // CString => int 변환,  단가
			str.Format(_T("%d"), m_pSet.db_m_Price);  // int => CString으로 변환
			m_listView2.SetItemText(0, 2, str);

			// 주문 갯수 및 합계 처리
			int n1 = 1;  // 기본값 갯수 1
			CString n1_string;
			n1_string.Format(_T("%d"), n1);

			m_listView2.SetItemText(0, 3, n1_string);  // 주문 갯수 리스트 컨트롤2에 출력

			OrderCount += n1; // 주문 갯수 합계 누적

			CString sum_string;
			sum_string.Format(_T("%d"), n1 * m_pSet.db_m_Price);
			m_listView2.SetItemText(0, 4, sum_string);

			OrderPrice += n1 * m_pSet.db_m_Price; // 주문 합계 누적
		}
	}

	// 기존 TotalCount에 주문 갯수를 추가
	TotalCount += OrderCount;

	// 전체 합계 출력
	strSum.Format(_T("%d"), OrderPrice);   // int => CString으로 변환
	SetDlgItemText(IDC_TOTAL_PRICE, strSum); // 전체 합계 표시

	// 주문 갯수 표시
	strCount.Format(TEXT("%d"), TotalCount);
	SetDlgItemText(IDC_TOTAL_COUNT, strCount);

	// 현재 시간 출력
	CTime gct = CTime::GetCurrentTime();
	CString strYear, strMonth, strDay, strTime, strYMDT;

	strYear.Format(_T("%d년 %d월 %d일"), gct.GetYear(), gct.GetMonth(), gct.GetDay());
	strTime.Format(_T(" %d시 %d분 %d초"), gct.GetHour(), gct.GetMinute(), gct.GetSecond());

	strYMDT = strYear + strTime;

	GetDlgItem(IDC_STATIC_TIME)->SetWindowText((LPCTSTR)strYMDT);  // 화면에 연월일 시간분초 출력
}




void CaaaaDlg::OnBnClickedMfcbuttonDel2()
{
	// 항목 삭제
	if (MessageBox(_T("현재 레코드를 삭제하시겠습니까?"), _T("레코드 삭제창"), MB_ICONQUESTION | MB_YESNO) == IDNO)
		return;

	for (int nItem2 = 0; nItem2 < m_listView2.GetItemCount();)
	{
		if (m_listView2.GetCheck(nItem2))
			m_listView2.DeleteItem(nItem2);
		else
			++nItem2;
	}

	AfxMessageBox(_T("레코드가 삭제 되었습니다."));
	UpdateData(FALSE); // 변수 => 컨트롤에 넣어 갱신(화면 출력)

	// 음식 주문 후 삭제된 가격 빼고, 카운트 다시 출력하기 시작
	int cnt = m_listView2.GetItemCount();
	CString strDel1;
	int TotalPrice = 0;

	for (int i = 0; i < cnt; i++)
	{
		if (m_listView2.GetCheck(i))
		{
			m_pSet.db_m_Price = _ttoi(m_listView2.GetItemText(i, 2)); // CString => int 변환
			strDel1.Format(_T("%d"), m_pSet.db_m_Price);              // int => CString으로 변환
			m_listView2.SetItemText(0, 2, strDel1);
		}

		// 주문한 음식 합계
		TotalPrice += _ttoi(m_listView2.GetItemText(i, 4)); // CString => int 변환
	}

	// 주문표(리스트 컨트롤2) 컨트롤 카운트
	int cnt2 = m_listView2.GetItemCount();
	CString strCount;
	strCount.Format(TEXT("%d"), cnt2);

	// 주문한 음식 합계 출력
	CString strSum;
	strSum.Format(_T("%d"), TotalPrice); // int => CString으로 변환
	SetDlgItemText(IDC_TOTAL_PRICE, strSum);

	// 주문 갯수 표시
	SetDlgItemText(IDC_TOTAL_COUNT, strCount);

	// 현재 시간 출력
	CTime gct = CTime::GetCurrentTime();
	CString strYear, strMonth, strDay, strTime, strYMDT;

	strYear.Format(_T("%d년 %d월 %d일"), gct.GetYear(), gct.GetMonth(), gct.GetDay());
	strTime.Format(_T(" %d시 %d분 %d초"), gct.GetHour(), gct.GetMinute(), gct.GetSecond());

	strYMDT = strYear + strTime;

	GetDlgItem(IDC_STATIC_TIME)->SetWindowText((LPCTSTR)strYMDT);  // 화면에 연월일 시간분초 출력
}


void CaaaaDlg::OnBnClickedMfcbuttonClear2()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	// 리스트 컨트롤2 화면 지우기와 관련된 부분 정리
	m_listView2.DeleteAllItems();

	// 주문 통계 정보 초기화
	TotalCount = 0;
	SetDlgItemText(IDC_TOTAL_COUNT, _T("0"));
	TotalPrice = 0;
	SetDlgItemText(IDC_TOTAL_PRICE, _T("0"));

	m_listTable.DeleteString(0);
	m_listTable.DeleteString(1);
}


void CaaaaDlg::OnBnClickedListprint()
{
	CPrintDialog dlgPrint(FALSE);

	if (dlgPrint.DoModal() == IDOK)
	{
		CDC dcPrint;
		dcPrint.Attach(dlgPrint.GetPrinterDC());

		DOCINFO docInfo;
		memset(&docInfo, 0, sizeof(DOCINFO));
		docInfo.cbSize = sizeof(DOCINFO);
		docInfo.lpszDocName = _T("ListView2 출력");
		docInfo.lpszOutput = NULL;
		docInfo.lpszDatatype = NULL;
		docInfo.fwType = 0;

		if (dcPrint.StartDoc(&docInfo) > 0)
		{
			if (dcPrint.StartPage() > 0)
			{
				CRect rectPrint(100, 100, 700, 1000);  // 출력 영역을 설정하세요.
				CRect rectPage(0, 0, dcPrint.GetDeviceCaps(HORZRES), dcPrint.GetDeviceCaps(VERTRES));

				CFont font;
				font.CreatePointFont(120, _T("Arial"));  // 폰트 및 크기를 조절하세요.
				dcPrint.SelectObject(&font);

				int colWidth = 200;  // 각 컬럼의 너비
				int rowHeight = 50;  // 각 행의 높이
				int textPadding = 5; // 텍스트와 셀 경계 간 여백

				// ListView2의 컬럼명을 출력합니다.
				for (int col = 0; col < m_listView2.GetHeaderCtrl()->GetItemCount(); col++)
				{
					HDITEM hdItem;
					ZeroMemory(&hdItem, sizeof(HDITEM));
					hdItem.mask = HDI_TEXT;
					hdItem.pszText = new TCHAR[MAX_PATH];
					hdItem.cchTextMax = MAX_PATH;

					if (m_listView2.GetHeaderCtrl()->GetItem(col, &hdItem))
					{
						CString strHeader = hdItem.pszText;

						// 컬럼명 출력
						dcPrint.TextOut(rectPrint.left + col * colWidth + textPadding, rectPrint.top + textPadding, strHeader);

						// 가로 선 그리기
						dcPrint.MoveTo(rectPrint.left + col * colWidth, rectPrint.top);
						dcPrint.LineTo(rectPrint.left + col * colWidth, rectPrint.bottom);
					}

					delete[] hdItem.pszText;
				}

				// ListView2의 각 항목을 출력합니다.
				for (int row = 0; row < m_listView2.GetItemCount(); row++)
				{
					for (int col = 0; col < m_listView2.GetHeaderCtrl()->GetItemCount(); col++)
					{
						CString strItem = m_listView2.GetItemText(row, col);

						// 항목 출력
						dcPrint.TextOut(rectPrint.left + col * colWidth + textPadding, rectPrint.top + (row + 1) * rowHeight + textPadding, strItem);

						// 세로 선 그리기
						dcPrint.MoveTo(rectPrint.left, rectPrint.top + (row + 1) * rowHeight);
						dcPrint.LineTo(rectPrint.right, rectPrint.top + (row + 1) * rowHeight);
					}
				}

				// total price 정보 출력
				CString strTotalPrice;
				GetDlgItemText(IDC_TOTAL_PRICE, strTotalPrice);
				CString strTotalPriceLabel = _T("결제 금액: ");
				dcPrint.TextOut(rectPrint.left, rectPrint.bottom + textPadding, strTotalPriceLabel);
				dcPrint.TextOut(rectPrint.left + 100, rectPrint.bottom + textPadding, strTotalPrice);

				dcPrint.EndPage();
			}

			dcPrint.EndDoc();
		}

		dcPrint.Detach();
	}
}





void CaaaaDlg::OnSelchangeCombo1()
{
	//---주문한 음식 테이블 번호-------------------
		int Index = m_comboTable.GetCurSel(); // 선택된 콤보 메뉴를 Index에 대입
	if (Index != CB_ERR)
	{
		//CString strTableNo;  // 헤더파일이 글로벌로 선언해야 화면에도 나타나고 종이에도 인쇄
		m_comboTable.GetLBText(Index, strTableNo);  // 주어진 항목의 문자열 조사
		m_comboTable.SetCurSel(Index);  // 주어진 항목을 선택 상태로 만듦​
		m_listTable.DeleteString(0);
		m_listTable.DeleteString(1);
		m_listTable.AddString(strTableNo); // 선택된 테이블 출력
	}
}

void CaaaaDlg::OnCbnSelchangeCombo2()
{
	int Index = m_comboCnt.GetCurSel();  // 콤보박스에서 선택된 값 가져오기

	int cnt2 = m_listView2.GetItemCount();
	int TotalPrice = 0;
	int TotalCount = _ttoi(strCount);  // 현재의 TotalCount를 가져와 초기값으로 사용

	for (int i = 0; i < cnt2; i++)
	{
		if (m_listView2.GetCheck(i) == TRUE)
		{
			// 주문 갯수 수정
			CString str2;
			str2.Format(_T("%d"), Index);
			m_listView2.SetItemText(i, 3, str2);

			// 주문 갯수에 따라 주문 금액 재계산
			int n2 = _ttoi(str2);
			m_pSet.db_m_Price = _ttoi(m_listView2.GetItemText(i, 2));
			CString sum_string;
			sum_string.Format(_T("%d"), n2 * m_pSet.db_m_Price);
			m_listView2.SetItemText(i, 4, sum_string);

			// TotalCount 업데이트
			TotalCount += n2;
		}
	}

	// 리스트 컨트롤2(음식주문표) 갯수 수정 후 다시 계산한 전체 합계와 카운트 출력
	for (int j = 0; j < cnt2; j++)
	{
		// 주문한 음식 합계 계산
		TotalPrice = TotalPrice + _ttoi(m_listView2.GetItemText(j, 4));
	}

	// 전체 합계 출력
	strSum.Format(_T("%d"), TotalPrice);
	SetDlgItemText(IDC_TOTAL_PRICE, strSum);

	// 전체 카운트 출력
	strCount.Format(TEXT("%d"), TotalCount-1);
	SetDlgItemText(IDC_TOTAL_COUNT, strCount);
}



void CaaaaDlg::OnBnClickedMfcbuttonExcel2()
{
	// 현재 시간 정보 가져오기
	CTime currentTime = CTime::GetCurrentTime();
	CString strYear, strTime;
	strYear.Format(_T("%d년 %d월 %d일"), currentTime.GetYear(), currentTime.GetMonth(), currentTime.GetDay());
	strTime.Format(_T(" %d시 %d분 %d초"), currentTime.GetHour(), currentTime.GetMinute(), currentTime.GetSecond());
	CString strYMDT = strYear + strTime;

	int m_iMax;
	int nColumncount = m_listView2.GetItemCount();
	int columnNum = 0;
	CString m_strFileName;

	m_iMax = nColumncount;

	CXLEzAutomation XL(FALSE); // 엑셀 API 함수를 사용하기 위한 클래스 변수 선언

	m_strFileName = strYMDT;

	XL.SetCellValue(++columnNum, 1, _T("코드"));
	XL.SetCellValue(++columnNum, 1, _T("메뉴"));
	XL.SetCellValue(++columnNum, 1, _T("가격"));
	XL.SetCellValue(++columnNum, 1, _T("갯수"));
	XL.SetCellValue(++columnNum, 1, _T("소계"));

	for (int i = 1; i <= m_iMax; i++)
	{
		XL.SetCellValue(1, i + 1, m_listview.GetItemText(i - 1, 0));

		for (int j = 1; j <= columnNum; j++)
			XL.SetCellValue(j + 1, i + 1, m_listView2.GetItemText(i - 1, j));
	}

	CFileDialog dlg(false, _T("xlsx"), _T("음식주문 백업") + m_strFileName + _T("_data.xlsx"),
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_NOCHANGEDIR,
		_T("xlsx 파일 (*.xlsx)|*.xlsx|)"), NULL);

	if (dlg.DoModal() == IDOK)
		XL.SaveFileAs(dlg.GetPathName());

	XL.ReleaseExcel();
}