# IIS 클래식 ASP OTP (2FA) 도입 구현 계획

이 계획은 IIS와 클래식 ASP 환경에서 로그인 보안을 강화하기 위해 Google OTP(Authenticator) 앱 기반의 **TOTP(Time-based One-Time Password)** 인증을 도입하는 방법에 대한 것입니다. 
C# .NET을 사용해 NuGet 의존성 없는 독립된 COM DLL을 제작하고, 이를 클래식 ASP에서 로드하여 사용합니다.

## User Review Required

> [!IMPORTANT]
> 1. **데이터베이스 변경 권장**: 
>    관리자 정보가 담겨있는 DB 테이블(예: `tblAdmin` 등)에 `AdmOtpSecret` (`VARCHAR(32) NULL`) 컬럼이 생성되어야 합니다.
>    (본 문서의 데이터베이스 스키마 수정안 쿼리를 실행해 주셔야 합니다.)
> 2. **IIS 서버 내 COM DLL 등록 작업**:
>    * 배포할 단독 실행형 파일은 **`OtpHelper.dll` 단 하나**입니다. (외부 의존성 파일 및 NuGet 패키지가 필요하지 않습니다.)
>    * 단순히 서버로 파일을 복사하는 것뿐 아니라, IIS 서버에서 관리자 권한으로 `RegAsm.exe`를 실행하여 레지스트리에 등록 작업을 진행해 주어야 합니다. (배포 가이드 참고)

---

## 데이터베이스 스키마 수정 권장안

관리자 테이블에 OTP 비밀키를 저장하기 위해 다음과 같이 컬럼 추가 및 초기 설정 쿼리를 수행해야 합니다. (실제 테이블명은 데이터베이스 환경에 맞추어 `tblAdmin` 또는 `Member_Admin` 등으로 변경해 주시기 바랍니다.)

```sql
-- 1. OTP 비밀키(Secret Key)를 저장할 컬럼 추가
ALTER TABLE tblAdmin ADD AdmOtpSecret VARCHAR(32) NULL;
GO
```

---

## Proposed Changes

### [Mpcjob.Otp] (C# COM DLL 생성)

외부 라이브러리(NuGet) 의존성이 없는 순수 .NET 4.0 호환 C# 클래스 라이브러리를 만듭니다.

#### [NEW] [OtpHelper.cs](file:///d:/MPCJOB/mpcjob_2012/admin/OtpHelper.cs)
* **내용**: Base32 디코더, Random Secret Key 생성, HMAC-SHA1 기반 TOTP(시간 동기 방식) 검증 기능 제공.
* **주요 메서드**:
  * `GenerateSecret()`: 16자리의 무작위 Base32 문자열 생성.
  * `VerifyOtp(secret, code)`: 현재 시간 전후 30초 오차 범위를 반영하여 6자리 OTP 입력값 검증.

---

### [Classic ASP 웹 서비스 연동 (테스트 분리 구성)]

기존 실서비스의 로그인/로그아웃 로직은 원상 보존하고, 검증을 테스트할 수 있도록 독립된 로그인 처리 파일들을 구성합니다.

#### [NEW] [login_otp.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login_otp.asp)
* **변경 사항**: 
  * 기존 `login.asp`를 복제하여 생성하였으며, ID/PW 검증 완료 후 `login_proc_otp.asp`를 타겟으로 하도록 양식의 action 속성을 수정하였습니다.

#### [NEW] [login_proc_otp.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc_otp.asp)
* **변경 사항**: 
  * 로그인 성공 시 임시 세션(`Session("TempASeq")` 등)을 부여하고, 데이터베이스에서 `AdmOtpSecret` 정보를 찾아 존재하면 `otp_auth.asp`로, 없으면 `otp_register.asp`로 리다이렉트합니다.

#### [MODIFY] [logout_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/logout_proc.asp)
* **변경 사항**: 로그아웃 시 임시 OTP 세션 변수들도 확실하게 지우도록 목록을 추가하되, 기존 환경에 미칠 영향을 차단하기 위해 **주석 처리(`'`)** 하였습니다.

#### [NEW] [otp_auth.asp](file:///d:/MPCJOB/mpcjob_2012/admin/otp_auth.asp)
* **역할**: 구글 OTP 앱의 6자리 번호를 입력받아 검증하는 화면 및 프로세스.
* **로직**:
  1. `Session("TempASeq")`가 존재하지 않으면 강제 로그인 아웃 처리 후 메인 로그인으로 리다이렉트.
  2. 사용자가 6자리 입력 시, C# COM 객체 `Mpcjob.Otp.OtpVerifier`를 생성하여 검증 수행.
  3. 성공 시 임시 세션을 실제 관리자 세션(`Session("ASeq")` 등)으로 변환 후 최종 관리자 홈으로 이동.

#### [NEW] [otp_register.asp](file:///d:/MPCJOB/mpcjob_2012/admin/otp_register.asp)
* **역할**: 최초 OTP 등록을 위한 QR 코드 표시 및 첫 검증 절차 페이지.
* **로직**:
  1. 신규 고유 Secret Key를 `GenerateSecret()`을 통해 생성.
  2. Google Chart API 또는 js 라이브러리를 통해 QR 코드(`otpauth://totp/MPCJOB:아이디?secret=비밀키&issuer=MPCJOB`) 화면 출력.
  3. 사용자가 앱에서 QR 스캔 후 나타나는 6자리 숫자를 입력하면 검증 처리.
  4. 검증 성공 시 DB의 관리자 레코드 `AdmOtpSecret` 필드에 생성된 Secret Key를 저장하고 최종 로그인 처리.

---

## Verification Plan

### 수동 검증 단계

1. **C# DLL 컴파일 및 IIS 등록**:
   * **관리자 권한 명령 프롬프트(cmd)**를 실행합니다.
   * `csc.exe`를 사용하여 `OtpHelper.cs`를 `OtpHelper.dll`로 빌드합니다:
     ```cmd
     C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:library /out:"D:\MPCJOB\mpcjob_2012\admin\OtpHelper.dll" "D:\MPCJOB\mpcjob_2012\admin\OtpHelper.cs"
     ```
   * `RegAsm.exe /codebase /tlb` 명령어로 IIS에 DLL을 COM 객체로 등록합니다. (IIS 응용 프로그램 풀 설정에 맞춰 64비트 혹은 32비트 등록 진행)
     * **64비트 IIS (기본값)**:
       ```cmd
       C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase /tlb "D:\MPCJOB\mpcjob_2012\admin\OtpHelper.dll"
       ```
     * **32비트 IIS (32비트 응용 프로그램 허용 시)**:
       ```cmd
       C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe /codebase /tlb "D:\MPCJOB\mpcjob_2012\admin\OtpHelper.dll"
       ```
     * *팁: 실제 비트수가 확실하지 않은 경우 두 명령어 모두 실행하는 것을 권장합니다.*
2. **최초 등록 테스트**:
   * 웹 브라우저로 `login_otp.asp`에 접속하여 로그인 시도.
   * OTP 미등록 사용자이므로 `otp_register.asp`로 이동하여 QR 코드와 비밀키가 정상 노출되는지 확인.
   * 스마트폰 구글 Authenticator 앱으로 QR 코드 스캔 후, 30초마다 갱신되는 6자리 번호 입력.
   * 인증 성공 시 DB에 비밀키가 반영되고 로그인이 통과하는지 확인.
3. **이후 로그인 테스트**:
   * 다시 `login_otp.asp`로 로그인 시도. DB에 키가 존재하므로 `otp_auth.asp`로 이동하는지 확인.
   * 임의의 6자리 오답 입력 시 로그인 실패 및 경고창 노출 확인.
   * 앱의 정답 6자리 입력 시 정상적으로 로그인이 최종 성공하는지 확인.
