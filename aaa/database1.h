#pragma once

#include <afxdb.h>

class database1 : public CRecordset
{
public:
    database1(CDatabase* pDatabase = NULL);
    DECLARE_DYNAMIC(database1)

    // �ʵ�/�Ű� ���� ������
    // �Ʒ��� ���ڿ� ����(���� ���)�� �����ͺ��̽� �ʵ��� ���� ������ ������
    // ��Ÿ���ϴ�(CStringA: ANSI ������ ����, CStringW: �����ڵ� ������ ����).
    // �̰��� ODBC ����̹����� ���ʿ��� ��ȯ�� ������ �� ������ �մϴ�.
    // ���� ��� �̵� ����� CString �������� ��ȯ�� �� ������
    // �׷� ��� ODBC ����̹����� ��� �ʿ��� ��ȯ�� �����մϴ�.
    // (����: �����ڵ�� �̵� ��ȯ�� ��� �����Ϸ��� ODBC ����̹�
    // ���� 3.5 �̻��� ����ؾ� �մϴ�).

    // �����ͺ��̽� �ʵ庰 ���̵�
    CString db_m_Code;      // ���̵�, �����ͺ��̽�
    CString db_m_Menu;    // �̸�
    int db_m_Price;       // �����ü�

    CString db_m_Image; // �̹��� ����

    // ������
    // �����翡�� ������ ���� �Լ� ������
public:
    virtual CString GetDefaultConnect(); // �⺻ ���� ���ڿ�
    virtual CString GetDefaultSQL();     // ���ڵ� ������ �⺻ SQL
    virtual void DoFieldExchange(CFieldExchange* pFX); // RFX ����

    // ����
#ifdef _DEBUG
    virtual void AssertValid() const;
    virtual void Dump(CDumpContext& dc) const;
#endif
};
