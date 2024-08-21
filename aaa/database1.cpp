#include "pch.h"
#include "database1.h"

IMPLEMENT_DYNAMIC(database1, CRecordset)

database1::database1(CDatabase* pdb) : CRecordset(pdb)
{
    db_m_Code = L"";
    db_m_Menu = L"";
    db_m_Price = 0;
    db_m_Image = L"";

    m_nFields = 4;
    m_nDefaultType = dynaset;
}

//#error 보안 문제: 연결 문자열에 암호가 포함되어 있을 수 있습니다.
// 아래 연결 문자열에 일반 텍스트 암호 및/또는 
// 다른 중요한 정보가 포함되어 있을 수 있습니다.
// 보안 관련 문제가 있는지 연결 문자열을 검토한 후에 #error을(를) 제거하십시오.
// 다른 형식으로 암호를 저장하거나 다른 사용자 인증을 사용하십시오.
CString database1::GetDefaultConnect()
{
    return _T("DSN=cafemenu");
}

CString database1::GetDefaultSQL()
{
    return _T("[table1]");
}

void database1::DoFieldExchange(CFieldExchange* pFX)
{
    pFX->SetFieldType(CFieldExchange::outputColumn);

    // RFX_Text() 및 RFX_Int() 같은 매크로는 데이터베이스의 필드
    // 형식이 아니라 멤버 변수의 형식에 따라 달라집니다.
    // ODBC에서는 자동으로 열 값을 요청된 형식으로 변환하려고 합니다
    RFX_Text(pFX, _T("[CODE]"), db_m_Code);
    RFX_Text(pFX, _T("[MENU]"), db_m_Menu);
    RFX_Int(pFX, _T("[PRICE]"), db_m_Price);
    RFX_Text(pFX, _T("[IMAGES]"), db_m_Image);
}

#ifdef _DEBUG
void database1::AssertValid() const
{
    CRecordset::AssertValid();
}

void database1::Dump(CDumpContext& dc) const
{
    CRecordset::Dump(dc);
}
#endif //_DEBUG
