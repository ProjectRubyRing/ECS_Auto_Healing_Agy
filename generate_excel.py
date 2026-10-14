import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_ecs_analysis_workbook():
    wb = openpyxl.Workbook()
    
    # 共通フォント定義 (Meiryo UI)
    FONT_NAME = "Meiryo UI"
    
    # スタイルパレット (洗練されたモノトーン基調)
    C_HEADER_BG = "212529"      # 最も濃いチャコールグレー
    C_SUBHEADER_BG = "495057"   # ダークグレー
    C_SECTION_BG = "6C757D"     # ミディアムグレー
    C_ZEBRA = "F8F9FA"          # ごく淡いグレー (偶数行用)
    C_CARD_BG = "F1F3F5"        # 情報ボックス背景
    C_BORDER = "CED4DA"         # 罫線用ライトグレー
    C_BORDER_DARK = "495057"    # 強調罫線用
    C_TEXT_MAIN = "212529"      # メイン文字色
    C_TEXT_MUTED = "6C757D"     # 補足文字色
    C_HIGHLIGHT = "343A40"      # 強調背景
    
    # フォント定義
    font_title = Font(name=FONT_NAME, size=15, bold=True, color="212529")
    font_subtitle = Font(name=FONT_NAME, size=10, italic=True, color="495057")
    font_section = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    font_tbl_header = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    font_bold = Font(name=FONT_NAME, size=10, bold=True, color=C_TEXT_MAIN)
    font_regular = Font(name=FONT_NAME, size=9.5, color=C_TEXT_MAIN)
    font_code = Font(name="Consolas", size=9, color="212529")
    font_badge = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
    
    # 罫線定義
    thin_side = Side(border_style="thin", color=C_BORDER)
    thick_dark_left = Side(border_style="medium", color=C_BORDER_DARK)
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_top_bottom = Border(top=thin_side, bottom=thin_side)
    border_card = Border(left=thick_dark_left, right=thin_side, top=thin_side, bottom=thin_side)
    
    # 塗りつぶし定義
    fill_header = PatternFill(start_color=C_HEADER_BG, end_color=C_HEADER_BG, fill_type="solid")
    fill_subheader = PatternFill(start_color=C_SUBHEADER_BG, end_color=C_SUBHEADER_BG, fill_type="solid")
    fill_section = PatternFill(start_color=C_SECTION_BG, end_color=C_SECTION_BG, fill_type="solid")
    fill_zebra = PatternFill(start_color=C_ZEBRA, end_color=C_ZEBRA, fill_type="solid")
    fill_card = PatternFill(start_color=C_CARD_BG, end_color=C_CARD_BG, fill_type="solid")
    fill_highlight = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")

    # 配置定義
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # -------------------------------------------------------------
    # Sheet 1: 事象総括と核心結論 (Executive Summary)
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "事象総括・核心結論"
    ws1.views.sheetView[0].showGridLines = True
    
    # タイトル
    ws1["A1"] = "【ECSオートヒーリング徹底調査】ヘルスチェック異常タスクの非停止と代替タスク停止の分析"
    ws1["A1"].font = font_title
    ws1["A2"] = "Desired Count = 6 における動作原理・仕様整合性・因果関係の完全解明レポート"
    ws1["A2"].font = font_subtitle
    
    # サマリーカードブロック (A4:F7)
    ws1.merge_cells("A4:F4")
    ws1["A4"] = "◆ 結論：これはECSの仕様上「完全に発生しうる事象」であり、ECSが正常に安全保護を働かせた結果です"
    ws1["A4"].font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    ws1["A4"].fill = fill_header
    ws1["A4"].alignment = align_left
    
    summary_text = (
        "【発生メカニズムの核心】\n"
        "1. 異常タスク（6個目）が停止しない理由：\n"
        "   ECSサービススケジューラーは「可用性最優先の安全原則（Rolling Replacement / minimumHealthyPercent）」に基づき、\n"
        "   新しく立ち上げる『代替タスク』が正常（HEALTHY & RUNNING）に稼働準備を整えるまで、異常な旧タスクを安易に強制終了（Kill）しません。\n"
        "2. 代替タスクが非アクティブ化から停止する理由：\n"
        "   代替タスク側で『起動時ヘルスチェック失敗』『ターゲットグループ(ALB)ヘルスチェック失敗』等の問題が発生し、起動直後に落とされています。\n"
        "3. 連鎖的膠着状態（デッドロック）：\n"
        "   代替タスクが死に続けるため『交代要員が永遠に来ない』状態となり、ECSは可用性を維持すべく異常タスクに停止命令（SIGTERM）を出せず、\n"
        "   結果として『異常タスクは起動し続け、裏で代替タスクが起動・停止を無限に繰り返す』という事象が発生します。"
    )
    ws1.merge_cells("A5:F8")
    ws1["A5"] = summary_text
    ws1["A5"].font = font_regular
    ws1["A5"].fill = fill_card
    ws1["A5"].alignment = align_top_left
    for r in range(5, 9):
        for c in range(1, 7):
            ws1.cell(row=r, column=c).border = border_card

    # パラメータ対比表
    ws1["A10"] = "■ 構成パラメータとシステム判定ステータス"
    ws1["A10"].font = font_bold

    headers_1 = ["パラメータ項目", "設定値 / 状態", "システム上の意味", "今回の影響とECSスケジューラーの挙動判断"]
    for col_idx, h in enumerate(headers_1, 1):
        cell = ws1.cell(row=11, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_subheader
        cell.alignment = align_header
        cell.border = border_all
    ws1.row_dimensions[11].height = 24

    params_data = [
        ("Desired Count (希望タスク数)", "6", "サービスが維持すべきタスク数", "常時6つのタスクを健全に稼働させる義務を負う"),
        ("バックコンテナの essential", "true", "タスク生存に不可欠なコンテナ", "このコンテナが死んだらタスク全体を停止・置換する対象となる"),
        ("コンテナヘルスチェック状態", "UNHEALTHY", "ls移動によりチェック失敗を検知", "ECSエージェントが異常を検知し、サービススケジューラーへ置換を要求"),
        ("現在稼働中の正常タスク数", "5 (タスク1〜5)", "健全と認められているタスク", "Desired Count(6) に対して 1つ 不足している状態"),
        ("minimumHealthyPercent", "100% (標準値)", "維持すべき最小健全率 (6×100%=6)", "健全タスクが5個しかないため、これ以上タスクを落とすと規律違反となる"),
        ("maximumPercent", "200% (標準値)", "起動可能な最大タスク率 (6×200%=12)", "新しい代替タスクを一時的に追加起動することは許可されている"),
        ("代替タスクの起動結果", "DEACTIVATING → STOPPED", "起動準備完了前に自滅・停止", "交代要員が完成しないため、異常タスクの停止処理が永久に保留される")
    ]

    for row_idx, data in enumerate(params_data, 12):
        ws1.row_dimensions[row_idx].height = 22
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx in [1, 2] else font_regular
            cell.alignment = align_center if col_idx == 2 else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # 因果関係ステップ表
    start_r = 21
    ws1.cell(row=start_r, column=1, value="■ なぜ異常タスクが停止しないのか？ 完全因果関係ステップ").font = font_bold
    
    headers_step = ["ステップ", "実行主体", "発生しているイベント", "内部ロジックと理由", "観測される症状"]
    for col_idx, h in enumerate(headers_step, 1):
        cell = ws1.cell(row=start_r+1, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_subheader
        cell.alignment = align_header
        cell.border = border_all
    ws1.row_dimensions[start_r+1].height = 24

    steps_data = [
        ("Step 1", "作業者 (ECS Exec)", "バックコンテナ内のlsコマンドを/tmpへ移動", "/mnt/logsの読み取りチェックコマンドが実行不能になりexit code 127/異常値を返す", "コンテナヘルスチェックがUNHEALTHYに遷移"),
        ("Step 2", "ECS Agent", "コンテナの異常ステータスをECSコントロールプレーンへ通知", "essential:true のコンテナが異常になったため、タスク置換が必要と報告", "マネジメントコンソール上でタスクがUNHEALTHY表示"),
        ("Step 3", "ECS Service", "Desired Count=6維持のため、代替タスク（7個目）のプロビジョニングを開始", "minimumHealthyPercent(100%)を守るため、新タスクが正常稼働するまで旧タスクは残す", "新しいタスクが PROVISIONING → PENDING → ACTIVATING になる"),
        ("Step 4", "代替タスク", "代替タスクが起動途中で異常終了（非アクティブ化へ遷移）", "マウント異常、初期ヘルスチェック失敗、ALB連携失敗等により生存不能と判定", "タスクが DEACTIVATING → STOPPING → STOPPED となり消滅"),
        ("Step 5", "ECS Service", "交代要員が死んだため、異常な6個目のタスクの停止命令（SIGTERM）を保留", "今6個目を殺すと健全タスクが5個のまま確定し、サービス停止リスクが高まるため", "【質問の事象】異常タスクが何十分経っても起動し続ける"),
        ("Step 6", "ECS Service", "再度、別の代替タスク（8個目）を起動しようとする", "Desired Count=6 を達成するまでスケジューラーは諦めずにリトライを継続する", "【質問の事象】何度も代替タスクが起動しては停止するループ")
    ]

    for row_idx, data in enumerate(steps_data, start_r+2):
        ws1.row_dimensions[row_idx].height = 28
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx == 1 else font_regular
            cell.alignment = align_center if col_idx in [1, 2] else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # -------------------------------------------------------------
    # Sheet 2: タスクライフサイクルと状態遷移 (Task Lifecycle)
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="ECSライフサイクルと非アクティブ化")
    ws2.views.sheetView[0].showGridLines = True

    ws2["A1"] = "【深掘り解析】ECSタスク状態遷移ステートマシンと「非アクティブ化」の正体"
    ws2["A1"].font = font_title
    ws2["A2"] = "タスクが PROVISIONING から DEACTIVATING、STOPPED に至る内部プロセスとトリガー"
    ws2["A2"].font = font_subtitle

    headers_lc = ["ステータス名", "フェーズ", "内部処理・実行されるアクション", "正常時の遷移先", "異常発生時の挙動（なぜDEACTIVATINGになるか？）"]
    for col_idx, h in enumerate(headers_lc, 1):
        cell = ws2.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_all
    ws2.row_dimensions[4].height = 26

    lc_data = [
        ("PROVISIONING", "起動準備", "ENI（Elastic Network Interface）の割り当て、VPC IPアドレスの確保、ボリュームのアタッチ", "PENDING", "IPアドレス枯渇、ENI割り当て制限超過、セキュリティグループ不整合で即停止"),
        ("PENDING", "コンテナ準備", "コンテナエージェントがECRからイメージをPull、設定・環境変数・Secretsの展開", "ACTIVATING", "イメージ取得失敗（ImagePullBackOff）、認証失敗、Secrets Manager権限不足で失敗"),
        ("ACTIVATING", "サービス参加", "ロードバランサー（ALB/NLB）のTarget Groupへ登録、サービスディスカバリ（Cloud Map）登録", "RUNNING", "ALB登録失敗、ヘルスチェック猶予期間（GracePeriod）超過、初期通信不能で失敗"),
        ("RUNNING", "本稼働", "コンテナプロセスが実行中。Dockerヘルスチェック（CMD-SHELL等）が定期実行される", "DEACTIVATING", "ヘルスチェック連続失敗（UNHEALTHY）、コンテナ内メインプロセスのクラッシュ終了"),
        ("DEACTIVATING", "切り離し待機", "ロードバランサーの登録解除（Draining/接続終了待ち）、サービスディスカバリ解除", "STOPPING", "【今回の症状】新タスクが初期判定でNGとみなされ、即座に切り離し処理へ直行している状態"),
        ("STOPPING", "終了シグナル", "コンテナプロセスへ SIGTERM シグナルを送信。StopTimeout（標準30秒）待機", "STOPPED", "プロセスが正常終了しない場合、SIGKILL により強制終了される"),
        ("STOPPED", "完全停止", "コンテナが停止し、割り当てられていたENIやボリューム等のAWSリソースが返却される", "終了（消滅）", "コンソールやCLIで stoppedReason や exitCode が確認可能となる")
    ]

    for row_idx, data in enumerate(lc_data, 5):
        ws2.row_dimensions[row_idx].height = 32
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx == 1 else font_regular
            cell.alignment = align_center if col_idx in [1, 2, 4] else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # DEACTIVATING が発生する具体的条件の解説ブロック
    ws2["A14"] = "■ 「非アクティブ化（DEACTIVATING）」が発生する3大トリガーの技術的詳細"
    ws2["A14"].font = font_bold

    headers_trig = ["トリガー区分", "発生メカニズム", "該当パラメータ / 設定箇所", "今回のケースにおける可能性と診断ポイント"]
    for col_idx, h in enumerate(headers_trig, 1):
        cell = ws2.cell(row=15, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_subheader
        cell.alignment = align_header
        cell.border = border_all
    ws2.row_dimensions[15].height = 24

    trig_data = [
        ("ターゲットグループ（ALB）の初期ヘルスチェック失敗",
         "タスク起動直後にALBからのヘルスチェックリクエストに応答できず、Unhealthyと判定されて即時Drainingに回された。",
         "healthCheckGracePeriodSeconds\n(デフォルト0秒だと即死しやすい)",
         "【極めて高確率】新タスクが立ち上がっても猶予期間が短すぎる場合、アプリの初期化完了前にALBが失格判定を下してDEACTIVATINGへ送られます。"),
        ("タスク定義自体のコンテナヘルスチェック失敗",
         "タスク定義に埋め込まれているヘルスチェックコマンド自体が、新タスク起動時に最初からエラー（exit code 1等）になっている。",
         "healthCheck.startPeriod\nhealthCheck.retries / timeout",
         "【高確率】もしマウントされている /mnt/logs が共有ストレージ（EFS等）で書き込み不可、あるいはヘルスチェックの初期猶予(startPeriod)が不足している場合。"),
        ("コンテナプロセスの初期起動エラー / クラッシュ",
         "ENTRYPOINTやCMDの実行直後に、設定ファイル不在、DB接続不可、環境変数不整合でプロセスがexit(1)等で自ら終了した。",
         "コンテナ定義の Command / Environment / DependsOn",
         "【中確率】バックコンテナがフロントコンテナの完全起動を待つ dependsOn 設定があり、その前提が崩れている場合に発生します。")
    ]

    for row_idx, data in enumerate(trig_data, 16):
        ws2.row_dimensions[row_idx].height = 40
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx == 1 else font_regular
            cell.alignment = align_center if col_idx == 1 else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # -------------------------------------------------------------
    # Sheet 3: 代替タスク異常停止の原因究明マトリクス (Root Cause Matrix)
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="代替タスク停止の原因分析")
    ws3.views.sheetView[0].showGridLines = True

    ws3["A1"] = "【原因究明マトリクス】なぜ代替タスクは起動直後に停止（DEACTIVATING）してしまうのか？"
    ws3["A1"].font = font_title
    ws3["A2"] = "現場で発生する可能性のある要因の網羅的洗い出しと検証手順"
    ws3["A2"].font = font_subtitle

    headers_rc = ["要因分類", "具体的な原因・障害シナリオ", "発生確率", "ECS停止理由 (stoppedReason) の特徴", "確認すべきAWSリソース / コマンド", "恒久是正アクション"]
    for col_idx, h in enumerate(headers_rc, 1):
        cell = ws3.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_all
    ws3.row_dimensions[4].height = 26

    rc_data = [
        ("① ELB猶予期間不足", "healthCheckGracePeriodSeconds が0秒または短すぎ、アプリ起動前にALBが異常判定した", "★★★★★\n(最頻出)", "Task failed ELB health checks in (target-group ...)", "aws ecs describe-services\nhealthCheckGracePeriodSecondsの秒数確認", "猶予期間をアプリの起動時間＋余裕（例: 60〜300秒）に延長設定する"),
        ("② コンテナチェック初期猶予不足", "コンテナヘルスチェックの startPeriod が未設定で、プロセス初期化中にチェックが走り即死した", "★★★★☆\n(高頻出)", "Task failed container health checks", "タスク定義の healthCheck.startPeriod 確認", "startPeriod を 30〜60秒程度設けて初期化完了を待つ設定にする"),
        ("③ /mnt/logs ボリューム依存", "マウント先（EFSやホストパス）の権限不足、ロック、容量枯渇、またはlsコマンド検証時の残留異常", "★★★★☆\n(要注意)", "Container exited with non-zero code / Health check failed", "EFSのアクセスポイント設定、マウント先ディレクトリの書き込み権限確認", "ボリュームマウントの整合性を担保し、ヘルスチェックコマンドが正常終了できるか確認"),
        ("④ ポート重複 (EC2起動タイプ)", "EC2インスタンス上で hostPort が静的固定されており、同一ホストに新タスクを配置できず衝突", "★★★☆☆\n(環境依存)", "was unable to place a task / port is already allocated", "タスク定義の portMappings (hostPort: 0 または awsvpc モードか確認)", "ネットワークモードを awsvpc に移行するか、hostPort を 0 (動的ポート) に変更"),
        ("⑤ サブネットIPアドレス枯渇", "awsvpcモードで、VPCサブネット内の空きプライベートIPアドレスが0になりENIが作れない", "★★☆☆☆\n(大規模環境)", "ResourceInitializationError: unable to reserve ENI", "aws ec2 describe-subnets\nAvailableIpAddressCount を確認", "サブネットの拡張、または空きIPのある別サブネットをサービス定義に追加"),
        ("⑥ デプロイサーキットブレーカー", "連続起動失敗により、ECSのDeployment Circuit Breakerが発動し自動ロールバック中", "★★☆☆☆\n(機能有効時)", "Deployment circuit breaker tripped", "aws ecs describe-services\ndeploymentCircuitBreaker のステータス確認", "失敗している根本タスクエラーを解消し、デプロイを再トリガーする")
    ]

    for row_idx, data in enumerate(rc_data, 5):
        ws3.row_dimensions[row_idx].height = 42
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws3.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx in [1, 3] else font_regular
            cell.alignment = align_center if col_idx in [1, 3] else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # -------------------------------------------------------------
    # Sheet 4: 歴史と背景・必要になった経緯 (History & Evolution)
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="登場の歴史と背景・必要性")
    ws4.views.sheetView[0].showGridLines = True

    ws4["A1"] = "【歴史的背景】なぜECSは「異常タスクを即座に殺さない」仕様に進化したのか？"
    ws4["A1"].font = font_title
    ws4["A2"] = "コンテナ黎明期のカスケード障害から、高可用性オーケストレーションへの変遷"
    ws4["A2"].font = font_subtitle

    headers_hist = ["年代・フェーズ", "技術的パラダイム", "当時の動作仕様", "直面した本番障害（なぜ問題だったか？）", "進化によって得られた恩恵"]
    for col_idx, h in enumerate(headers_hist, 1):
        cell = ws4.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_all
    ws4.row_dimensions[4].height = 26

    hist_data = [
        ("第1世代\n(2014-2015年)\nDocker初期・原始的コンテナ管理",
         "プロセス即時Killモデル",
         "ヘルスチェックが異常になったら即座にSIGKILL等でコンテナを強制終了し、ゼロから立ち上げ直す。",
         "【カスケード障害の嵐】\n一時的な高負荷でタイムアウトしたコンテナが即死。残ったコンテナに負荷が集中し、次々とドミノ倒しのように全コンテナが全滅した。",
         "シンプルな構造であったが、大規模な本番Webサービスでは全断リスクが極めて高かった。"),
        ("第2世代\n(2016-2018年)\nECS Rolling Update & Minimum Healthy",
         "可用性最優先・\nローリング交代モデル",
         "Desired CountとminimumHealthyPercentを導入。新タスクがReadyになるまで旧タスクを生かす基本方針を確立。",
         "【交代難民問題】\n新タスクの定義にミスがある場合、新タスクが死に続け、旧タスクが異常のままゾンビ化する（今回直面している現象の原型）。",
         "一時的なスパイクによる全コンテナ全滅を完全に防げるようになり、システムの高可用性が劇的に向上した。"),
        ("第3世代\n(2019-2022年)\nContainer Health Check統合",
         "多層的ヘルスチェック連携モデル",
         "ALBだけでなく、Dockerコンテナ内部のCMD-SHELLヘルスチェック結果をECSスケジューラーが直接評価可能に。",
         "【多層判定の複雑化】\n「コンテナヘルスチェック」「ALBターゲットグループチェック」「タスクステータス」の3層が食い違い、なぜ止まらないか混乱が生じるようになった。",
         "Webサーバー以外のバックグラウンドワーカーでも、コンテナ内部の異常を自動検知してオートヒーリングが可能になった。"),
        ("第4世代\n(2023年-現在)\nサーキットブレーカー & ガバナンス",
         "自律的保護と異常ループ遮断モデル",
         "デプロイメントサーキットブレーカーやタスク起動失敗の抑制機能が強化され、無限ループを自動検出・ロールバック。",
         "【ブラックボックス化】\n安全機構が多重化されたため、現場のエンジニアが仕様を知らないと「バグでコンテナが止まらない」と誤認してしまう。",
         "人間が介入しなくても、誤ったデプロイによるタスク連続死から自動復旧できる究極の自律保護を実現。")
    ]

    for row_idx, data in enumerate(hist_data, 5):
        ws4.row_dimensions[row_idx].height = 48
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws4.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx in [1, 2] else font_regular
            cell.alignment = align_center if col_idx in [1, 2] else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # -------------------------------------------------------------
    # Sheet 5: 小学生でもわかる図解ストーリー (Kid-friendly Analogy)
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="小学生でもわかる図解ストーリー")
    ws5.views.sheetView[0].showGridLines = True

    ws5["A1"] = "【小学生でも絶対にわかる】大人気ハンバーガーショップの「6番レジ係」の悲劇"
    ws5["A1"].font = font_title
    ws5["A2"] = "難しいクラウドの仕組みを、身近なハンバーガー屋さんの出来事に例えて完全解説！"
    ws5["A2"].font = font_subtitle

    # キャラクター対応表
    ws5["A4"] = "■ 登場人物とクラウド用語の対応表"
    ws5["A4"].font = font_bold

    headers_ch = ["お店のキャラクター", "実際のAWS / ECS用語", "どんなお仕事・役割？"]
    for col_idx, h in enumerate(headers_ch, 1):
        cell = ws5.cell(row=5, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_subheader
        cell.alignment = align_header
        cell.border = border_all
    ws5.row_dimensions[5].height = 24

    char_data = [
        ("店長さん", "ECSサービススケジューラー", "お店全体を見張って、「レジは絶対に6台開けておけ！」と命令するボス"),
        ("6台のレジ係 (1番〜6番)", "タスク定義 (Desired Count = 6)", "お客さんの注文を受けてハンバーガーを渡す6人の店員さん"),
        ("6番レジの店員さん", "今回エラーになった6個目のタスク", "フライ返し（lsコマンド）を失くして、お腹が痛くなってしまった店員さん"),
        ("フライ返しを/tmpに移動", "lsコマンドの移動 (疑似障害テスト)", "道具を隠して、お仕事ができないように意図的にイジワルしたこと"),
        ("「お腹が痛いです！」の申告", "コンテナヘルスチェック UNHEALTHY", "店長さんに「私、もう正常に仕事ができません！」とSOSを出した状態"),
        ("交代のアルバイトさん", "新しく起動しようとする代替タスク", "店長さんが電話で呼び出した、助っ人の新しい店員さん"),
        ("お店のルール (最低維持率)", "minimumHealthyPercent (100%)", "「絶対に6台のレジを動かせ。交代が来るまで絶対に帰るな！」という絶対掟")
    ]

    for row_idx, data in enumerate(char_data, 6):
        ws5.row_dimensions[row_idx].height = 24
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws5.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx in [1, 2] else font_regular
            cell.alignment = align_center if col_idx == 1 else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # ストーリー展開
    ws5["A15"] = "■ 事件の真相：なぜ6番レジの店員さんは帰れないのか？（4コマ解説）"
    ws5["A15"].font = font_bold

    headers_st = ["場面", "お店で起きていること (ストーリー)", "クラウド(ECS)の中で起きていること", "店長(ECS)の心の声"]
    for col_idx, h in enumerate(headers_st, 1):
        cell = ws5.cell(row=16, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_all
    ws5.row_dimensions[16].height = 24

    story_data = [
        ("第1話\n【腹痛発生】",
         "6番レジの店員さんがフライ返しを失くし、「お腹が痛くて動けません…」と座り込んでしまった！",
         "コンテナ内のlsコマンドが無くなり、定期監視のヘルスチェックがエラー（UNHEALTHY）になった。",
         "「大変だ！でも今6番を帰らせたらレジが5台になってお店が大混乱になる！急いで交代のバイトを呼ぼう！」"),
        ("第2話\n【交代要員の召集】",
         "店長さんは、事務所から交代の新しいアルバイトさん（代替タスク）を呼び出した！",
         "ECSは Desired Count=6 を守るため、新しい代替タスクを立ち上げ始めた（ACTIVATINGへ遷移）。",
         "「よし、この新しいバイトがエプロンを着てレジに入ったら、お腹の痛い6番を病院へ帰らせてあげよう」"),
        ("第3話\n【バイトのドタキャン】",
         "ところが！やってきたアルバイトさんは、お店の鍵が開けられず、準備中に転んで「無理です」と帰ってしまった！",
         "代替タスクが起動時のチェックやALBの通信チェックに合格できず、非アクティブ化して停止（STOPPED）した。",
         "「おいおい！バイトが途中で帰っちゃったぞ！？ これじゃ6番を帰らせるわけには絶対にいかない！」"),
        ("第4話\n【永遠の居残り】",
         "店長さんは何人も何人も新しいバイトを呼ぶが、全員途中で帰ってしまう！ 6番レジの店員さんは痛みに耐えて立ち続ける…",
         "代替タスクが起動失敗を繰り返すため、ECSは異常タスクを停止（Kill）できず、異常なまま起動させ続けている！",
         "「交代がちゃんとレジに立つまで、お前は絶対にそこを動くな！ （＝タスクを停止させない安全装置）」")
    ]

    for row_idx, data in enumerate(story_data, 17):
        ws5.row_dimensions[row_idx].height = 44
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws5.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx == 1 else font_regular
            cell.alignment = align_center if col_idx == 1 else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # -------------------------------------------------------------
    # Sheet 6: トラブルシューティング手順とコマンド (Runbook)
    # -------------------------------------------------------------
    ws6 = wb.create_sheet(title="調査・トラブルシューティング手順")
    ws6.views.sheetView[0].showGridLines = True

    ws6["A1"] = "【現場用ランブック】代替タスク停止の真因を突き止める5分間切り分け手順"
    ws6["A1"].font = font_title
    ws6["A2"] = "AWS CLI と CloudWatch / マネジメントコンソールを用いた確定診断フロー"
    ws6["A2"].font = font_subtitle

    headers_rb = ["No", "調査対象・確認項目", "実行する具体的なAWS CLIコマンド / 操作", "確認すべき出力フィールド", "判断基準と次のアクション"]
    for col_idx, h in enumerate(headers_rb, 1):
        cell = ws6.cell(row=4, column=col_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_all
    ws6.row_dimensions[4].height = 26

    rb_data = [
        ("Step 1", "停止した代替タスクの特定",
         "aws ecs list-tasks --cluster <ClusterName> --service-name <ServiceName> --desired-status STOPPED --max-items 5",
         "taskArns (直近で停止した代替タスクのARNリスト)",
         "最新のタスクARNをコピーし、Step 2のコマンドに渡す"),
        ("Step 2", "停止理由 (stoppedReason) の取得",
         "aws ecs describe-tasks --cluster <ClusterName> --tasks <TaskArn> --query \"tasks[].[taskArn, stoppedReason, stopCode, containers[].[name, exitCode, reason]]\"",
         "stoppedReason\nstopCode\nexitCode",
         "『failed container health checks』ならStep 3へ\n『failed ELB health checks』ならStep 4へ"),
        ("Step 3", "コンテナヘルスチェックの猶予確認",
         "aws ecs describe-task-definition --task-definition <TaskDefName> --query \"taskDefinition.containerDefinitions[].[name, healthCheck]\"",
         "healthCheck.startPeriod\nhealthCheck.retries\nhealthCheck.command",
         "startPeriod が 0 または未設定なら、アプリ初期化完了を待てる秒数（例: 60秒）を設定する"),
        ("Step 4", "ALBヘルスチェック猶予期間の確認",
         "aws ecs describe-services --cluster <ClusterName> --services <ServiceName> --query \"services[].[serviceName, healthCheckGracePeriodSeconds]\"",
         "healthCheckGracePeriodSeconds",
         "0秒になっている場合は、Target Groupのヘルスチェック間隔×回数以上の猶予（例: 120〜300秒）に更新する"),
        ("Step 5", "ECSサービスイベントの履歴確認",
         "aws ecs describe-services --cluster <ClusterName> --services <ServiceName> --query \"services[0].events[0:10].[createdAt, message]\" --output table",
         "events[].message",
         "『unable to place a task』等のリソース枯渇やポート競合メッセージが出ていないかを確認する")
    ]

    for row_idx, data in enumerate(rb_data, 5):
        ws6.row_dimensions[row_idx].height = 36
        fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
        for col_idx, val in enumerate(data, 1):
            cell = ws6.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_code if col_idx in [3, 4] else (font_bold if col_idx == 1 else font_regular)
            cell.alignment = align_center if col_idx == 1 else align_left
            cell.border = border_all
            if fill.fill_type:
                cell.fill = fill

    # -------------------------------------------------------------
    # 全シート共通：列幅の自動調整・ビュー設定・固定枠
    # -------------------------------------------------------------
    for ws in wb.worksheets:
        # ウィンドウ枠の固定 (タイトル行固定)
        ws.freeze_panes = "A5"
        
        # 列幅の手動最適化設定（Meiryo UIのフォントサイズに合わせて視認性を最大化）
        if ws.title == "事象総括・核心結論":
            ws.column_dimensions["A"].width = 28
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 45
            ws.column_dimensions["D"].width = 45
            ws.column_dimensions["E"].width = 30
            ws.column_dimensions["F"].width = 25
        elif ws.title == "ECSライフサイクルと非アクティブ化":
            ws.column_dimensions["A"].width = 18
            ws.column_dimensions["B"].width = 16
            ws.column_dimensions["C"].width = 45
            ws.column_dimensions["D"].width = 20
            ws.column_dimensions["E"].width = 50
        elif ws.title == "代替タスク停止の原因分析":
            ws.column_dimensions["A"].width = 22
            ws.column_dimensions["B"].width = 38
            ws.column_dimensions["C"].width = 16
            ws.column_dimensions["D"].width = 35
            ws.column_dimensions["E"].width = 38
            ws.column_dimensions["F"].width = 38
        elif ws.title == "登場の歴史と背景・必要性":
            ws.column_dimensions["A"].width = 20
            ws.column_dimensions["B"].width = 24
            ws.column_dimensions["C"].width = 38
            ws.column_dimensions["D"].width = 42
            ws.column_dimensions["E"].width = 38
        elif ws.title == "小学生でもわかる図解ストーリー":
            ws.column_dimensions["A"].width = 20
            ws.column_dimensions["B"].width = 42
            ws.column_dimensions["C"].width = 42
            ws.column_dimensions["D"].width = 42
        elif ws.title == "調査・トラブルシューティング手順":
            ws.column_dimensions["A"].width = 12
            ws.column_dimensions["B"].width = 26
            ws.column_dimensions["C"].width = 50
            ws.column_dimensions["D"].width = 30
            ws.column_dimensions["E"].width = 45

    output_path = r"C:\Users\taka_\Agy\ECS_Auth_Healing_Agy\ECS_Auto_Healing_Analysis.xlsx"
    wb.save(output_path)
    print(f"Workbook successfully saved to: {output_path}")

if __name__ == "__main__":
    create_ecs_analysis_workbook()
