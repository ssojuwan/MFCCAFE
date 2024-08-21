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

//#error ���� ����: ���� ���ڿ��� ��ȣ�� ���ԵǾ� ���� �� �ֽ��ϴ�.
// �Ʒ� ���� ���ڿ��� �Ϲ� �ؽ�Ʈ ��ȣ ��/�Ǵ� 
// �ٸ� �߿��� ������ ���ԵǾ� ���� �� �ֽ��ϴ�.
// ���� ���� ������ �ִ��� ���� ���ڿ��� ������ �Ŀ� #error��(��) �����Ͻʽÿ�.
// �ٸ� �������� ��ȣ�� �����ϰų� �ٸ� ����� ������ ����Ͻʽÿ�.
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

    // RFX_Text() �� RFX_Int() ���� ��ũ�δ� �����ͺ��̽��� �ʵ�
    // ������ �ƴ϶� ��� ������ ���Ŀ� ���� �޶����ϴ�.
    // ODBC������ �ڵ����� �� ���� ��û�� �������� ��ȯ�Ϸ��� �մϴ�
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
