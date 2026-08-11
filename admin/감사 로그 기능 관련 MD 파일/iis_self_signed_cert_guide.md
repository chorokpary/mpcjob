# Windows Server 2016 IIS - test.mpcjob.co.kr 사설 SSL 인증서 생성 및 적용 가이드

본 문서는 Windows Server 2016 (IIS 10.0) 환경에서 개발 및 테스트용 도메인인 `test.mpcjob.co.kr`에 사설(Self-Signed) SSL/TLS 인증서를 생성하고, IIS 웹 사이트에 바인딩(HTTPS) 및 클라이언트 PC 신뢰 등록까지 적용하는 전체 과정을 정리한 가이드입니다.

---

## 📌 1. 개요 및 사전 준비

* **운영체제 및 웹서버**: Windows Server 2016 / IIS 10.0
* **대상 도메인**: `test.mpcjob.co.kr`
* **주요 목적**: 개발/테스트 환경에서 보안 프로토콜(HTTPS) 및 감사 로그 / SSL 세션 테스트 적용
* **주의사항**: IIS 내 기본 "자체 서명된 인증서" 기능은 SAN(Subject Alternative Name) 속성이 누락되어 최신 웹 브라우저(Chrome, Edge 등)에서 `NET::ERR_CERT_COMMON_NAME_INVALID` 오류가 발생할 수 있습니다. 따라서 **PowerShell의 `New-SelfSignedCertificate` cmdlet을 활용하여 SAN 속성을 포함한 인증서를 생성**하는 것을 권장합니다.
* **기존 IIS 서비스 영향도**: 
  - **영향 전혀 없음 (서비스 중단 0초)**
  - IIS 서비스 재시작(`iisreset`)이나 앱풀(AppPool) 재회가 필요하지 않으며, 기존에 운영 중인 HTTP/HTTPS 사이트 접속에 아무런 영향을 주지 않습니다.
  - (포트 충돌 방지: 기존에 포트 443을 사용하는 다른 SSL 사이트가 있을 경우 포트 충돌을 방지하기 위해 본 가이드에서는 HTTPS 포트로 **`4443`**(또는 다른 여유 포트)을 사용하여 바인딩합니다.)

---

## 🛠️ 2. 1단계: PowerShell을 이용한 SAN 사설 인증서 생성

관리자 권한으로 **Windows PowerShell**을 실행한 후 아래 명령어를 입력하여 인증서를 생성합니다.

### ① 인증서 생성 명령어 (PowerShell 관리자 권한)

```powershell
# 1. test.mpcjob.co.kr 사설 SSL 인증서 생성 (유효기간 10년)
New-SelfSignedCertificate -DnsName "test.mpcjob.co.kr", "localhost" -CertStoreLocation "cert:\LocalMachine\My" -NotAfter (Get-Date).AddYears(10) -FriendlyName "test.mpcjob.co.kr Self-Signed Cert"

### ② 생성 결과 확인
명령어 실행 후 반환되는 `Thumbprint`(지문) 및 `Subject` 정보를 확인합니다.
```text
Thumbprint                                Subject
----------                                -------
A1B2C3D4E5F678901234567890ABCDEF12345678  CN=test.mpcjob.co.kr
```
> 해당 인증서는 Windows Server의 `개인(Personal)` 인증서 저장소(`cert:\LocalMachine\My`)에 자동으로 저장되며, IIS 관리자에서 즉시 선택할 수 있는 상태가 됩니다.

---

## ⚙️ 3. 2단계: IIS 신규 웹 사이트 생성 및 HTTPS 바인딩 설정

기존 운영 사이트와 독립된 테스트 환경을 구축하기 위해 IIS에서 **신규 애플리케이션 풀 및 웹 사이트를 생성**하고 SSL 바인딩을 추가합니다.

### ① 신규 애플리케이션 풀(App Pool) 생성
1. `Windows 키 + R` $\rightarrow$ `inetmgr` 입력하여 IIS 관리자 실행
2. 좌측 연결 창에서 **`애플리케이션 풀 (Application Pools)`** 우클릭 $\rightarrow$ **`애플리케이션 풀 추가...`** 선택
3. 항목 설정:
   - **이름**: `AppPool_test_mpcjob`
   - **.NET CLR 버전**: `.NET CLR 버전 없음` (Classic ASP 전용인 경우) 또는 `.NET CLR v4.0...`
   - **관리되는 파이프라인 모드**: `통합 (Integrated)`
4. `확인` 클릭

### ② 신규 웹 사이트 생성 및 HTTPS 바인딩
1. 좌측 연결 창에서 **`사이트 (Sites)`** 우클릭 $\rightarrow$ **`웹 사이트 추가... (Add Website...)`** 선택
2. [웹 사이트 추가] 창에서 아래 항목 설정:
   - **사이트 이름**: `test.mpcjob.co.kr`
   - **애플리케이션 풀**: [선택] 클릭하여 위에서 생성한 `AppPool_test_mpcjob` 지정
   - **실제 경로 (Physical Path)**: 소스코드 위치 지정 (예: `D:\MPCJOB\mpcjob_2012`)
    - **[바인딩 설정]**:
      * **종류 (Type)**: `https`
      * **IP 주소 (IP address)**: `모두 미지정` (또는 특정 IP)
      * **포트 (Port)**: `4443` (다른 사이트와의 443 포트 충돌을 피하기 위해 `4443` 또는 임의의 여유 포트를 지정합니다)
      * **호스트 이름 (Host name)**: `test.mpcjob.co.kr`
      * **`SNI(서버 이름 표시) 필요`**: **체크 (Checked)** (포트 충돌 방지 및 여러 SSL 바인딩 관리를 위해 체크 권장)
      * **SSL 인증서**: 1단계에서 생성한 **`test.mpcjob.co.kr Self-Signed Cert`** 선택
3. `확인` 클릭하여 사이트 생성 완료

### ③ Classic ASP 설정 (Classic ASP 소스 사용 시)
1. 생성된 `test.mpcjob.co.kr` 사이트 선택 후 중앙 기능 뷰에서 **`ASP`** 기능 더블 클릭
2. **`부모 경로 사용 (Enable Parent Paths)`**: `True`로 변경
3. 우측 [작업] 패널에서 `적용` 클릭

---

## 🌐 4. 3단계: 로컬 DNS (Hosts 파일) 설정 (선택 사항 - DNS 미연결 시에만 진행)

> **💡 이미 DNS(공인/내부 도메인 서버)에 `test.mpcjob.co.kr`이 해당 서버 IP로 연결되어 있다면 이 단계는 생략(Skip)하시면 됩니다.**  
> DNS 레코드가 등록되어 있지 않은 개발/테스트 초기 환경이거나, 로컬 PC에서 임시로 도메인을 서버 IP로 강제 매핑해야 하는 경우에만 아래 작업을 진행합니다.

1. **메모장(Notepad)을 관리자 권한으로 실행**
2. **경로 파일 열기**: `C:\Windows\System32\drivers\etc\hosts`
3. **하단에 매핑 정보 추가**:
   ```text
   # test.mpcjob.co.kr 바인딩 테스트
   127.0.0.1       test.mpcjob.co.kr
   # (외부 테스트 PC인 경우 서버의 실제 내부/외부 IP 주소 입력)
   # 192.168.1.100   test.mpcjob.co.kr
   ```
4. 저장 후 종료

---

## 🔒 5. 4단계: 클라이언트 PC 인증서 신뢰 등록 (경고 페이지 해제)

사설 인증서는 공인 인증 기관(CA)이 서명한 것이 아니므로, 최초 접속 시 웹 브라우저에서 **"사용자의 연결은 비공개로 설정되어 있지 않습니다 (`NET::ERR_CERT_AUTHORITY_INVALID`)"** 경고가 출력됩니다.  
클라이언트 PC(또는 서버)의 `신뢰할 수 있는 루트 인증 기관`에 인증서를 등록하여 빨간색 경고를 해제합니다.

### ① 인증서 내보내기 (.cer 파일)

PowerShell(관리자 권한)에서 내보내기 명령을 실행합니다:

```powershell
# C:\certs 폴더 생성 및 인증서 내보내기
New-Item -ItemType Directory -Path "C:\certs" -Force
Get-ChildItem -Path "cert:\LocalMachine\My" | Where-Object { $_.FriendlyName -eq "test.mpcjob.co.kr Self-Signed Cert" } | Export-Certificate -FilePath "C:\certs\test.mpcjob.co.kr.cer" -Force
```

### ② 신뢰할 수 있는 루트 인증 기관에 가져오기

PowerShell(관리자 권한)에서 내보낸 인증서를 `Root` 저장소에 등록합니다:

```powershell
# 신뢰할 수 있는 루트 인증 기관으로 인증서 가져오기
Import-Certificate -FilePath "C:\certs\test.mpcjob.co.kr.cer" -CertStoreLocation "Cert:\LocalMachine\Root"
```

> **GUI 방식으로 등록하는 방법**:
> 1. `C:\certs\test.mpcjob.co.kr.cer` 파일 더블 클릭
> 2. `인증서 설치(I)...` 클릭 $\rightarrow$ 저장소 위치: `로컬 머신(Local Machine)` 선택
> 3. `모든 인증서를 다음 저장소에 저장` 선택 후 `신뢰할 수 있는 루트 인증 기관` 지정
> 4. 마침 및 설치 승인

---

## 🧪 6. 5단계: 최종 접속 검증 및 문제 해결 (Troubleshooting)

### ① 웹 브라우저 접속 검증
1. 웹 브라우저(Chrome / Edge)를 열고 `https://test.mpcjob.co.kr:4443` 으로 접속합니다. (지정한 포트번호 `:4443`을 주소 뒤에 반드시 붙여야 합니다.)
2. 주소창 왼쪽에 **자물쇠 아이콘(🔒)**이 정상 표시되며 `이 연결은 안전합니다` 상태인지 확인합니다.

---

### ② 주요 이슈 해결 (Troubleshooting)

| 발생 증상 | 원인 | 조치 방법 |
| :--- | :--- | :--- |
| **`NET::ERR_CERT_COMMON_NAME_INVALID`** | 인증서에 `SAN` 속성(DNS Name)이 없음 | 2단계의 PowerShell 스크립트(`-DnsName "test.mpcjob.co.kr"`)로 인증서를 재생성 |
| **`NET::ERR_CERT_AUTHORITY_INVALID`** | 클라이언트 PC가 사설 CA를 신뢰하지 않음 | 5단계의 `Import-Certificate`를 통해 `LocalMachine\Root`에 인증서 등록 |
| **사이트 접속 불가 (시간 초과)** | Windows 방화벽 설정(포트 4443 등) 차단 | `인바운드 규칙`에서 **사용한 포트 (예: TCP 4443)** 허용 추가 |
| **Classic ASP 세션/쿠키 분리** | HTTP $\leftrightarrow$ HTTPS 간 세션 끊김 | `web.config` 또는 쿠키 설정에 `httpOnly`, `sameSite` 및 `Secure` 쿠키 설정 점검 |
