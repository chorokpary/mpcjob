# OTP (2FA) 구현 관련 생성 및 수정 파일 리스트

IIS 클래식 ASP 환경에 구글/마이크로소프트 OTP 추가 인증(2FA)을 적용하기 위해 생성 및 수정된 파일들의 전체 목록입니다.

---

## 1. 🆕 새로 생성한 파일

| 구분 | 파일 경로 | 설명 |
| :--- | :--- | :--- |
| **테스트 로그인 화면** | [login_otp.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login_otp.asp) | 기존 로그인 양식을 기반으로, OTP 인증 처리(`login_proc_otp.asp`) 페이지를 바라보는 테스트 화면입니다. |
| **테스트 로그인 처리** | [login_proc_otp.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc_otp.asp) | 로그인 성공 후 유저의 OTP 연동 정보(Secret Key)가 존재하면 `otp_auth.asp`로, 없다면 `otp_register.asp`로 리다이렉트하는 페이지입니다. |
| **최초 OTP 등록** | [otp_register.asp](file:///d:/MPCJOB/mpcjob_2012/admin/otp_register.asp) | 유저의 DB 정보에 OTP 비밀키가 없을 때, 구글 OTP 연동 QR 코드를 제공하고 검증 성공 시 DB에 비밀키를 최종 등록합니다. |
| **OTP 인증 번호 입력** | [otp_auth.asp](file:///d:/MPCJOB/mpcjob_2012/admin/otp_auth.asp) | 기기 연동이 완료된 유저가 로그인 후 6자리 일회용 비밀번호를 입력하고 통과하는 2차 인증 화면입니다. |
| **핵심 계산 모듈 (C#)** | [OtpHelper.cs](file:///d:/MPCJOB/mpcjob_2012/admin/OtpHelper.cs) | Base32 디코딩 및 HMAC-SHA1 기반 표준 TOTP 알고리즘 계산 로직을 가진 C# 클래스 라이브러리 소스코드입니다. |
| **핵심 라이브러리 (DLL)** | [OtpHelper.dll](file:///d:/MPCJOB/mpcjob_2012/admin/OtpHelper.dll) | `csc.exe` 컴파일러로 빌드 완료한 COM 등록용 DLL 어셈블리 파일입니다. |

---

## 2. ✏️ 기존 소스에서 수정한 파일

| 구분 | 파일 경로 | 설명 |
| :--- | :--- | :--- |
| **관리자 로그아웃 처리** | [logout_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/logout_proc.asp) | 로그아웃 시 임시 세션들도 일괄 제거하는 변수 정리를 추가했습니다. 기존 운영 환경 영향 방지를 위해 해당 부분은 **주석 처리(`'`)**되었습니다. (추후 실서비스 배포 시 주석 해제 요망) |

---

## 3. ↩️ 원상 복구 완료 파일 (운영 환경 유지)

테스트 및 운영 안정을 위해 이전 수정 이력을 완전히 롤백하여 **원본 그대로 보존**된 파일들입니다.

* [login.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login.asp): 수정 내역 없음 (원본 상태 유지).
* [login_proc.asp](file:///d:/MPCJOB/mpcjob_2012/admin/login_proc.asp): 추가했던 OTP 분기 처리 로직을 완전히 제거하고, 원래의 1단계 즉시 로그인 통과 방식으로 롤백 완료.
