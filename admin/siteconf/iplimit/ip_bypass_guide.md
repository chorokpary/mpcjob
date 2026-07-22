# IP 접근 제한 예외 관리자 ID 관리 기능 적용 가이드

본 문서는 관리자 접속 IP 접근 제한 시스템에 특정 관리자 ID에 대한 예외 허용(우회) 기능을 팝업 형태로 동적 관리할 수 있도록 구성된 구현 내용 및 설명서입니다.

---

## 1. 🗄️ 필수 데이터베이스 작업 (SQL 쿼리)

예외 관리자 ID 목록을 동적으로 추가 및 삭제하기 위해 데이터베이스(MSSQL)에 아래 테이블을 생성해 주셔야 합니다.

```sql
-- 예외 관리자 ID 저장 테이블 생성 쿼리
CREATE TABLE TBL_IP_BYPASS (
    BypassSeq INT IDENTITY(1,1) PRIMARY KEY,
    AdminId VARCHAR(50) NOT NULL,
    BypassMemo VARCHAR(200) NULL,
    RegDate DATETIME DEFAULT GETDATE()
);
```

---

## 2. 📂 작업 및 수정 파일 목록

| 구분 | 파일 경로 | 설명 |
| :--- | :--- | :--- |
| **IP 관리 메인** | [ip_list.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_list.asp) | 상단 영역에 `🔓 예외 ID 관리` 버튼 추가 및 팝업창(`openBypassPop()`) 호출 스크립트 반영 |
| **예외 ID 팝업** | [ip_bypass_pop.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_bypass_pop.asp) | 등록된 예외 관리자 ID 목록 조회, 신규 예외 ID 및 메모 입력 폼, 삭제 버튼 제공 팝업 |
| **팝업 처리 로직** | [ip_bypass_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_bypass_proc.asp) | 예외 ID 등록(`insert`) 및 삭제(`delete`) DB 쿼리 수행 모듈 |
| **IP 접근 검증** | [ip_check.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_check.asp) | 로그인한 계정이 `TBL_IP_BYPASS` 테이블에 존재하면 IP 차단 대상에서 무조건 예외 처리(통과) |

---

## 3. 🔄 주요 기능 및 프로세스

### Step 1. 예외 ID 관리 팝업 호출
* 관리자 페이지의 IP 접근 제한 설정([ip_list.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_list.asp)) 접속 후 `내 IP 자동입력` 옆의 **`🔓 예외 ID 관리`** 버튼을 클릭합니다.
* `ip_bypass_pop.asp` 팝업창(크기: 620x520)이 새로 열립니다.

### Step 2. 예외 ID 등록 및 삭제
* **등록**: 팝업 상단 입력 폼에 예외 처리할 관리자 ID(예: `admin_test`)와 식별용 메모를 입력 후 등록 버튼을 누릅니다.
* **삭제**: 목록에서 삭제할 ID 항목의 `삭제` 버튼을 누르면 DB에서 제거되어 예외 처리가 해제됩니다.

### Step 3. IP 검증 우회 (자동 적용)
* `TBL_IP_BYPASS`에 등록된 아이디로 로그인한 사용자는 **[ip_check.asp](file:///d:/MPCJOB/mpcjob_2012/admin/siteconf/iplimit/ip_check.asp)** 접근 제어 체크 시 **IP 허용 목록에 등록되어 있지 않더라도 무조건 정상 통과**합니다.
