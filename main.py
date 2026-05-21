import win32com.client
import datetime
import os


def get_and_merge_todays_ppts():
    # 1. 다운로드 및 저장 경로 설정
    current_dir = os.path.abspath(os.getcwd())
    download_folder = os.path.join(current_dir, "회의자료_다운로드")
    merged_file_path = os.path.join(current_dir, "주간보고병합.pptx")

    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    # 부서 및 담당자 목록 (자료취합순서 기준)
    team_info = [
        {"order": 1, "team": "연구기획팀", "manager": "허용민"},
        {"order": 2, "team": "심혈관팀", "manager": "김태원"},
        {"order": 3, "team": "급성감염팀", "manager": "한예지"},
        {"order": 4, "team": "Cancer팀", "manager": "정진용"},
        {"order": 5, "team": "호르몬팀", "manager": "이소희"},
        {"order": 6, "team": "치료용항체팀", "manager": "김영은"},
        {"order": 7, "team": "갑상선팀", "manager": "김세희"},
        {"order": 8, "team": "당뇨팀", "manager": "함은선"}
    ]

    # 제출 여부 추적용 딕셔너리
    # {담당자이름: False} 형태로 초기화
    submitted_status = {info["manager"]: False for info in team_info}

    # 병합할 파일들을 순서와 함께 저장할 리스트
    downloaded_files_with_order = []

    try:
        # 2. 아웃룩 연결 및 설정
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        today = datetime.date.today()
        today_str = today.strftime("%y%m%d")  # 오늘 날짜 (예: 260521)

        print(f"📅 오늘 날짜 검색 형식: '{today_str}'")
        print("📥 메일 검색 및 PPT 다운로드를 시작합니다...\n" + "=" * 60)

        # 최신 메일부터 확인
        for message in messages:
            try:
                recv_date = datetime.date(message.ReceivedTime.year,
                                          message.ReceivedTime.month,
                                          message.ReceivedTime.day)

                if recv_date < today:
                    break  # 어제 날짜로 넘어가면 탐색 종료
                if recv_date > today:
                    continue

                subject = message.Subject if message.Subject else ""
                sender = message.SenderName if message.SenderName else ""
                attachments = message.Attachments

                if attachments.Count == 0:
                    continue

                has_target_ppt = False
                matched_attachments = []

                # 첨부파일 분석 (PPT + 날짜포함)
                for att in attachments:
                    filename = att.FileName
                    ext = filename.lower()
                    if (ext.endswith('.ppt') or ext.endswith('.pptx')) and (today_str in filename):
                        has_target_ppt = True
                        matched_attachments.append(att)

                if has_target_ppt:
                    # 메일 보낸 사람이 담당자 목록에 있는지 확인
                    matched_order = 99  # 명단에 없는 사람일 경우 기본 순서값을 뒤로 뺌
                    matched_team = "알수없음"

                    for info in team_info:
                        if info["manager"] in sender:
                            matched_order = info["order"]
                            matched_team = info["team"]
                            submitted_status[info["manager"]] = True
                            break

                    print(f"[{matched_team}] 담당자 '{sender}'의 메일 확인됨: {subject}")

                    # 파일 다운로드
                    for att in matched_attachments:
                        time_str = message.ReceivedTime.strftime("%H%M%S")
                        safe_filename = f"{matched_team}_{time_str}_{att.FileName}"
                        save_path = os.path.join(download_folder, safe_filename)

                        att.SaveAsFile(save_path)
                        # 순서와 파일 경로를 튜플로 묶어서 저장
                        downloaded_files_with_order.append((matched_order, save_path))
                        print(f"    -> 다운로드: {safe_filename}")

            except AttributeError:
                continue
            except Exception as e:
                continue

        # ==========================================
        # 3. 미제출 부서 확인 및 출력 (오류 수정됨)
        # ==========================================
        print("\n" + "=" * 60)
        print("📊 주간보고서 제출 현황")
        print("=" * 60)

        missing_teams = []
        for info in team_info:
            # 제출 상태가 False(안 냄)인 경우 missing_teams 리스트에 추가
            if not submitted_status[info["manager"]]:
                missing_teams.append(f"{info['team']}({info['manager']})")

        if missing_teams:
            print("❌ 미제출 부서 (총 {}곳): \n   -> {}".format(len(missing_teams), ", ".join(missing_teams)))
        else:
            print("✅ 모든 부서가 제출을 완료했습니다!")

        if not downloaded_files_with_order:
            print("\n다운로드된 취합 파일이 없습니다. 병합을 종료합니다.")
            return

        # ==========================================
        # 4. 파워포인트 병합 (취합 순서대로 정렬)
        # ==========================================
        print("\n" + "=" * 60)
        print("🔄 지정된 취합 순서대로 PPT를 병합합니다...")

        # 저장된 리스트를 order(순서) 기준으로 오름차순 정렬
        downloaded_files_with_order.sort(key=lambda x: x[0])
        sorted_files = [file_path for order, file_path in downloaded_files_with_order]

        ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        try:
            # 첫 번째 파일(순서가 가장 빠른 파일)을 베이스로 열기
            base_ppt = ppt_app.Presentations.Open(sorted_files[0], WithWindow=False)

            # 두 번째 파일부터 베이스 파일 끝에 순서대로 병합
            for file_path in sorted_files[1:]:
                current_slide_count = base_ppt.Slides.Count
                base_ppt.Slides.InsertFromFile(file_path, current_slide_count)
                print(f"    -> 병합 추가 완료: {os.path.basename(file_path)}")

            # '주간보고병합.pptx' 로 저장
            base_ppt.SaveAs(merged_file_path)
            base_ppt.Close()

            print("\n✅ 병합이 완료되었습니다! 📂 파일명: 주간보고병합.pptx")

        except Exception as e:
            print(f"PPT 병합 중 오류 발생: {e}")
        finally:
            ppt_app.Quit()

    except Exception as e:
        print(f"프로그램 실행 중 치명적 오류 발생: {e}")


if __name__ == "__main__":
    get_and_merge_todays_ppts()
