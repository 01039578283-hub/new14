# 배포 전 검증

GitHub 저장소는 `01039578283-hub/new14`, 운영 브랜치는 `main`, Vercel 프로젝트는 `new14`입니다. 작업 중인 원본 폴더 대신 검토된 변경만 있는 작업 트리에서 배포합니다.

## 한 번 설치

Node.js 24와 Git을 사용합니다. 저장소를 새로 받았거나 아래 검사 도구를 변경했다면 실행합니다.

```sh
node tools/release/install-pre-push.mjs
```

이 명령은 해당 로컬 저장소의 공통 Git 디렉터리에 검사 도구를 설치합니다. 연결된 모든 worktree에서 적용되며 특정 작업 폴더의 존속에 의존하지 않습니다. 기존 `pre-push`는 원본 그대로 백업하고 검사 성공 후 같은 인자·표준입력으로 실행합니다. Git LFS 업로드 훅도 유지됩니다. 별도의 `core.hooksPath` 설정이 있으면 덮어쓰지 않고 중단합니다.

## 설정 또는 빌드 파일 변경 시

1. `git diff`로 의도한 변경을 검토합니다. `vercel.json`의 리다이렉트·헤더·빌드 명령·출력 폴더도 배포 입력입니다.
2. 변경을 확인한 파일 하나씩만 검증 목록에 기록합니다. 아래 명령은 페이지·이미지 목록을 수정하지 않으며 다른 파일의 변경을 일괄 승인하지 않습니다.

```sh
node release-preflight.mjs --review-input vercel.json
node release-preflight.mjs
node --test tests/release-preflight.test.mjs tests/kakao-directions.test.mjs
node kakao-directions-build.mjs
```

3. 변경 파일과 `release-public-manifest.json`을 함께 커밋합니다. 기존 URL·본문·이미지에 의도하지 않은 변경이 없는지도 검토합니다.
4. 실제 보낼 커밋을 검사한 뒤 명시적 대상에 푸시합니다.

```sh
node release-preflight.mjs --ref HEAD
git push origin HEAD:main
```

훅은 작업 폴더나 stage 상태가 아니라 푸시로 전달되는 각 브랜치의 **커밋 내용**을 읽습니다. 작업 폴더에서만 검증 목록을 고쳐 놓고 커밋하지 않은 경우에도 푸시는 중단됩니다. 검증기·빌드·훅은 해시를 자동 갱신하지 않습니다. 새 빌드 입력을 추가하면 `release-inputs.mjs`의 입력 목록과 검증 목록을 함께 검토하고 훅을 다시 설치합니다. 예전 형식의 커밋을 롤백할 때도 현재 입력 계약에 맞춘 별도 검토가 필요합니다.

## 환경 차이와 실패 처리

- 텍스트의 LF/CRLF 차이와 Vercel 설정 JSON의 공백·객체 키 순서는 허용합니다.
- 설정 값, 배열 순서, 리다이렉트 추가·삭제, 실제 코드 변경은 검토 기록 없이는 차단합니다.
- 오류 메시지에 표시된 파일의 실제 변경을 확인합니다. 실패를 없애려고 자동 승인, `--no-verify`, 검증 제거를 사용하지 않습니다.
- GitHub Actions는 Windows와 Linux에서 회귀 검사·커밋 검사·전체 공개 빌드를 실행합니다. 원격 CI는 푸시 이후 실행되므로 Vercel 자동 배포보다 먼저 완료되는 장치가 아닙니다. 로컬 사전 차단은 설치된 훅이 담당합니다.
- GitHub 웹 편집, 훅이 없는 다른 PC, 명시적인 훅 우회에는 로컬 차단이 적용되지 않습니다. 이 경로에서도 CI와 Vercel 빌드 검사는 유지되지만 실패 알림을 완전히 방지한다고 보장할 수 없습니다.

## 배포 완료 확인

Vercel의 `READY`, 운영 도메인의 배포 연결, GitHub 커밋 SHA 일치, 실제 공개 HTML·이미지 응답을 따로 확인합니다. 미리보기 실패·취소와 운영 배포 실패를 구분합니다. 홈페이지 내용이 바뀌지 않은 배포 검사 수정이라면 이전 공개 산출물과의 동일성도 검사합니다.

참고: [Git pre-push 문서](https://git-scm.com/docs/githooks#_pre_push), [Vercel GitHub 배포](https://vercel.com/docs/git/vercel-for-github).
