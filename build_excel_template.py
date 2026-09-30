# -*- coding: utf-8 -*-
import os
import shutil
import xlsxwriter

def create_system_architecture_template():
    filename = "AWS_現場システム統合構成図_テンプレート.xlsx"
    workbook = xlsxwriter.Workbook(filename, {'default_date_format': 'yyyy/mm/dd'})
    
    # 0. Global Font Setting - Strictly Meiryo UI
    workbook.formats[0].set_font_name('Meiryo UI')
    workbook.formats[0].set_font_size(9)
    
    # -------------------------------------------------------------
    # 1. FORMAT DEFINITIONS (Monotone Palette & Meiryo UI)
    # -------------------------------------------------------------
    def fmt(base_dict):
        d = {'font_name': 'Meiryo UI', 'font_size': 9}
        d.update(base_dict)
        return workbook.add_format(d)
        
    # Headers & Titles
    f_sheet_title = fmt({'font_size': 13, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1A202C', 'valign': 'vcenter', 'indent': 1})
    f_sheet_subtitle = fmt({'font_size': 9, 'font_color': '#CBD5E0', 'bg_color': '#1A202C', 'valign': 'vcenter', 'align': 'right'})
    f_sec_title = fmt({'font_size': 10.5, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#2D3748', 'valign': 'vcenter', 'indent': 1})
    f_sec_sub = fmt({'font_size': 9, 'bold': True, 'font_color': '#2D3748', 'bg_color': '#EDF2F7', 'valign': 'vcenter', 'indent': 1, 'bottom': 1, 'bottom_color': '#4A5568'})

    # Meta Block Formats
    f_meta_label = fmt({'font_size': 8.5, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#4A5568', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#718096'})
    f_meta_val = fmt({'font_size': 8.5, 'font_color': '#1A202C', 'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E0'})

    # Table Formats
    f_tbl_head = fmt({'font_size': 8.5, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#2D3748', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#4A5568', 'text_wrap': True})
    f_tbl_cell = fmt({'font_size': 8.5, 'font_color': '#1A202C', 'bg_color': '#FFFFFF', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E0', 'text_wrap': True})
    f_tbl_cell_center = fmt({'font_size': 8.5, 'font_color': '#1A202C', 'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E0', 'text_wrap': True})
    f_tbl_cell_code = fmt({'font_size': 8.5, 'font_name': 'Meiryo UI', 'font_color': '#1A202C', 'bg_color': '#F7FAFC', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E0'})
    f_tbl_cell_zebra = fmt({'font_size': 8.5, 'font_color': '#1A202C', 'bg_color': '#F7FAFC', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E0', 'text_wrap': True})
    f_tbl_cell_zebra_center = fmt({'font_size': 8.5, 'font_color': '#1A202C', 'bg_color': '#F7FAFC', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E0', 'text_wrap': True})

    # Container / Box Header Formats
    f_box_head_dark = fmt({'font_size': 9.5, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1A202C', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#1A202C', 'text_wrap': True})
    f_box_head_slate = fmt({'font_size': 9, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#2D3748', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#2D3748', 'text_wrap': True})
    f_box_head_gray = fmt({'font_size': 8.5, 'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#4A5568', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#4A5568', 'text_wrap': True})
    f_box_head_light = fmt({'font_size': 8.5, 'bold': True, 'font_color': '#1A202C', 'bg_color': '#EDF2F7', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#718096', 'text_wrap': True})

    # Container Body Fills
    f_bg_cloud = fmt({'bg_color': '#F7FAFC', 'border': 1, 'border_color': '#718096'})
    f_bg_vpc = fmt({'bg_color': '#EDF2F7', 'border': 1, 'border_color': '#4A5568'})
    f_bg_subnet_pub = fmt({'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#718096'})
    f_bg_subnet_pri = fmt({'bg_color': '#EDF2F7', 'border': 1, 'border_color': '#4A5568'})
    f_bg_subnet_db = fmt({'bg_color': '#E2E8F0', 'border': 1, 'border_color': '#2D3748'})
    f_bg_onprem = fmt({'bg_color': '#F7FAFC', 'border': 1, 'border_color': '#4A5568'})
    f_bg_white = fmt({'bg_color': '#FFFFFF', 'border': 1, 'border_color': '#CBD5E0'})

    # Note / Legend Formats
    f_note = fmt({'font_size': 8, 'font_color': '#4A5568', 'valign': 'vcenter', 'text_wrap': True})
    f_flow_arrow = fmt({'font_size': 11, 'bold': True, 'font_color': '#2D3748', 'align': 'center', 'valign': 'vcenter'})

    print("Formats initialized successfully.")

    # -------------------------------------------------------------
    # HELPER: Draw Container Box with Cell Borders
    # -------------------------------------------------------------
    def draw_cell_container(ws, r1, c1, r2, c2, title, head_format, body_format, head_rows=1):
        if head_rows == 1:
            ws.merge_range(r1, c1, r1, c2, title, head_format)
            start_body = r1 + 1
        else:
            ws.merge_range(r1, c1, r1 + head_rows - 1, c2, title, head_format)
            start_body = r1 + head_rows
            
        for r in range(start_body, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.write_blank(r, c, None, body_format)

    # -------------------------------------------------------------
    # HELPER: Write Standard Metadata Header
    # -------------------------------------------------------------
    def write_meta_header(ws, sheet_title, sub_desc, sys_name="現場IoT・基幹データ連携ハイブリッド統合基盤"):
        # Main Title Banner (Row 1, Cols B to AO)
        ws.merge_range(1, 1, 1, 28, sheet_title, f_sheet_title)
        ws.merge_range(1, 29, 1, 41, sub_desc, f_sheet_subtitle)
        ws.set_row(1, 28)
        
        # Meta Table (Row 2-3, Cols B to AO)
        ws.set_row(2, 18)
        ws.set_row(3, 18)
        
        # Row 2
        ws.merge_range(2, 1, 2, 3, "システム名称", f_meta_label)
        ws.merge_range(2, 4, 2, 15, sys_name, f_meta_val)
        ws.merge_range(2, 16, 2, 18, "設計バージョン", f_meta_label)
        ws.merge_range(2, 19, 2, 22, "Ver 1.0.0 (正式版)", f_meta_val)
        ws.merge_range(2, 23, 2, 25, "作成日", f_meta_label)
        ws.merge_range(2, 26, 2, 29, "2026/10/01", f_meta_val)
        ws.merge_range(2, 30, 2, 32, "機密区分", f_meta_label)
        ws.merge_range(2, 33, 2, 35, "社内限り (Confidential)", f_meta_val)
        ws.merge_range(2, 36, 2, 38, "作成者", f_meta_label)
        ws.merge_range(2, 39, 2, 41, "インフラ統括設計部", f_meta_val)

        # Row 3 (Approval block)
        ws.merge_range(3, 1, 3, 3, "対象リージョン", f_meta_label)
        ws.merge_range(3, 4, 3, 15, "AWS 東京リージョン (ap-northeast-1) / 大阪 (DR)", f_meta_val)
        ws.merge_range(3, 16, 3, 18, "更新日", f_meta_label)
        ws.merge_range(3, 19, 3, 22, "2026/10/01", f_meta_val)
        ws.merge_range(3, 23, 3, 25, "承認ステータス", f_meta_label)
        ws.merge_range(3, 26, 3, 29, "設計承認済 (Approved)", f_meta_val)
        ws.merge_range(3, 30, 3, 32, "承認印", f_meta_label)
        ws.write(3, 33, "【承認】", f_meta_label)
        ws.write(3, 34, "部長", f_meta_val)
        ws.write(3, 35, "【審査】", f_meta_label)
        ws.write(3, 36, "課長", f_meta_val)
        ws.write(3, 37, "【担当】", f_meta_label)
        ws.merge_range(3, 38, 3, 41, "アーキテクト", f_meta_val)


    # =========================================================================
    # SHEET 0: 00_利用ガイド・凡例規約
    # =========================================================================
    ws0 = workbook.add_worksheet('00_利用ガイド・凡例規約')
    ws0.hide_gridlines(0)
    ws0.set_landscape()
    ws0.set_paper(9) # A4
    ws0.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws0.set_column('A:A', 3)
    ws0.set_column('B:B', 18)
    ws0.set_column('C:C', 20)
    ws0.set_column('D:D', 26)
    ws0.set_column('E:E', 32)
    ws0.set_column('F:F', 24)
    ws0.set_column('G:G', 16)
    
    write_meta_header(ws0, "【利用ガイド】現場×AWS 統合システム構成図 テンプレート標準規約", "Meiryo UI / モノトーン設計標準")
    
    row = 5
    ws0.merge_range(row, 1, row, 6, "1. 本テンプレートの構成とシートの役割分担", f_sec_title)
    row += 1
    headers_s0 = ["シート名", "対象読者・スコープ", "表現内容・詳細度", "含まれる主な要素", "活用シーン"]
    for ci, h in enumerate(headers_s0):
        ws0.write(row, ci + 1, h, f_tbl_head)
    row += 1
    
    sheet_roles = [
        ("01_構成図_概要", "経営層・PM・全体関係者", "マクロ鳥瞰図 (1枚で全体像把握)", "現場設備群、拠点、専用線/VPN、AWS主要層、外部連携、全体仕様表", "提案書、プロジェクト全体説明、報告資料、システム俯瞰"),
        ("02_構成図_詳細", "インフラ・NW・開発・運用", "精密設計図 (実装・構築レベル)", "Multi-AZ VPC、サブネットCIDR、IP、ポート、SG、OT/IT分離、詳細機器一覧表", "基本設計書、詳細設計書、NW申請、セキュリティ監査、構築保守"),
        ("03_パーツ集_概要用", "構成図作成者", "概要図用コピペ部品カタログ", "マクロ境界枠、主要階層カード、データフローマクロ線、エグゼクティブ凡例", "概要図を新規作成・拡張カスタマイズする際の貼り付け元"),
        ("04_パーツ集_詳細用", "構成図作成者", "詳細図用コピペ部品カタログ", "AWS全70+アイコン、現場機器、詳細線種、ポートタグ、ステータスバッジ、表部品", "詳細図で大量のシステムやサービスを精密配置する際の貼り付け元"),
    ]
    for r_data in sheet_roles:
        ws0.write(row, 1, r_data[0], f_tbl_cell_center)
        ws0.write(row, 2, r_data[1], f_tbl_cell)
        ws0.write(row, 3, r_data[2], f_tbl_cell)
        ws0.write(row, 4, r_data[3], f_tbl_cell)
        ws0.write(row, 5, r_data[4], f_tbl_cell)
        ws0.set_row(row, 22)
        row += 1
        
    row += 1
    ws0.merge_range(row, 1, row, 6, "2. デザイン原則・カラーパレット定義 (モノトーンベース)", f_sec_title)
    row += 1
    
    headers_color = ["カラーコード", "カラー名称", "色彩イメージ", "主な適用対象", "設計意図・効果"]
    for ci, h in enumerate(headers_color):
        ws0.write(row, ci + 1, h, f_tbl_head)
    row += 1
    
    colors_info = [
        ("#1A202C", "Charcoal Dark", "極暗濃灰 (黒に近いスレート)", "シートメインタイトル、最重要境界、全体枠外枠", "視覚的な引き締めと最重要レベルの階層識別"),
        ("#2D3748", "Deep Slate", "濃スレートグレー", "大見出し、アイコン上部ヘッダー、主要境界線", "高コントラストで可読性に優れた基調ダークトーン"),
        ("#4A5568", "Medium Slate", "中間スレートグレー", "サブネット境界、二重線、第2階層見出し", "主要線と補助線の明確な視覚的差別化"),
        ("#718096", "Neutral Gray", "ニュートラルグレー", "非同期破線、管理点線、注記、セル区切り線", "情報量を増やしても画面がうるさくならない抑制トーン"),
        ("#E2E8F0", "Soft Gray", "薄灰塗りつぶし", "DBサブネット、重要カードヘッダー、表ゼブラ", "背景に自然になじむ落ち着いたゾーン強調"),
        ("#EDF2F7", "Light Slate", "極薄スレート塗りつぶし", "VPC背景、入力欄、マクロカード背景", "白背景と明確に区別できるコンテナ下地"),
        ("#F7FAFC", "Off-White", "クリーンオフホワイト", "AWS Cloud全体枠背景、現場プラント背景", "清潔感と高い視認性を両立する広域キャンバス"),
        ("#FFFFFF", "Pure White", "純白", "個別コンポーネントカード、アイコン背景", "印刷・PDF化・モノクロコピー時でも最高の明瞭度"),
    ]
    for c_code, c_name, c_img, c_target, c_intent in colors_info:
        c_fmt = workbook.add_format({'font_name': 'Meiryo UI', 'font_size': 8.5, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E0', 'bg_color': c_code, 'font_color': '#FFFFFF' if c_code in ['#1A202C', '#2D3748', '#4A5568'] else '#1A202C'})
        ws0.write(row, 1, c_code, c_fmt)
        ws0.write(row, 2, c_name, f_tbl_cell_center)
        ws0.write(row, 3, c_img, f_tbl_cell)
        ws0.write(row, 4, c_target, f_tbl_cell)
        ws0.write(row, 5, c_intent, f_tbl_cell)
        ws0.set_row(row, 20)
        row += 1

    row += 1
    ws0.merge_range(row, 1, row, 6, "3. 作図を極めて美しく仕上げる「セル方眼」操作テクニック", f_sec_title)
    row += 1
    
    tips = [
        ("【重要】Altキーを押しながらドラッグ (グリッドスナップ)", "Excel上で図形や貼り付けたアイコン画像を移動・リサイズする際、[Alt] キーを押しながらドラッグすると、自動的にセルの境界線（方眼格子）にピタリと吸着します。これにより、複数コンポーネントの配置ズレや不揃いが完全に防げます。"),
        ("Meiryo UI フォントの徹底", "Meiryo UI は日本語の行間余白が均一でコンパクトなため、狭い枠内でも文字が上下に切れず、極めて整然と表示されます。図形内テキストや注記を追加する際も、必ず Meiryo UI を維持してください。"),
        ("階層構造（Zオーダー）の意識", "作図時は「最背面：リージョン・VPC・現場エリアの背景セル」→「中間面：サブネット枠・接続線・通信フロー矢印」→「最前面：アイコン画像・テキストタグ」の順で重ねることで、整理された見やすい図面になります。"),
        ("大量のシステムを記載する場合のベストプラクティス", "現場の設備やAWSサービスが膨大になる場合、1枚の図面にすべてを詰め込まず、【01_概要】で全体ランドスケープ（マクロ接続）を提示し、【02_詳細】シートを業務ドメインやシステム系統ごとに複製（例：02_詳細_工場ライン系、02_詳細_物流EC系）して詳細化することを推奨します。"),
        ("印刷・PDF出力時の用紙設定", "概要図・詳細図ともに標準で「A3横向き / 横1ページに収める（Fit to 1 page wide）」設定を行っています。A4用紙に印刷・配布する場合も、Excelの印刷画面で「用紙サイズ：A4」を選択するだけで自動的に高品質縮小出力されます。")
    ]
    for title, desc in tips:
        ws0.write(row, 1, title, f_tbl_head)
        ws0.merge_range(row, 2, row, 6, desc, f_tbl_cell)
        ws0.set_row(row, 30)
        row += 1


    # =========================================================================
    # SHEET 1: 01_構成図_概要 (Overview Architecture Diagram Template)
    # =========================================================================
    ws1 = workbook.add_worksheet('01_構成図_概要')
    ws1.hide_gridlines(0)
    ws1.set_landscape()
    ws1.set_paper(8) # A3
    ws1.fit_to_pages(1, 0)
    ws1.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    
    # Configure micro-grid columns (A to AO: 41 columns)
    ws1.set_column('A:A', 2)
    for c in range(1, 42):
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        ws1.set_column(f'{col_letter}:{col_letter}', 3.4)
    for r in range(4, 70):
        ws1.set_row(r, 18)

    write_meta_header(ws1, "【全体俯瞰】現場設備・エッジ・拠点 ⇔ AWS クラウド 統合システム構成図 (概要)", "エグゼクティブ・PM向け 全体鳥瞰図")

    # ---------------- Pillar 1: On-premise / Edge ----------------
    draw_cell_container(ws1, 5, 1, 31, 10, "【現場・拠点領域】工場プラント・エッジ・事業所", f_box_head_dark, f_bg_onprem, head_rows=1)
    
    # Sub-container 1A: スマート工場プラント (OT系)
    draw_cell_container(ws1, 7, 2, 17, 9, "スマート製造工場 (OT系ライン設備)", f_box_head_slate, f_bg_white)
    ws1.write(8, 2, "PLC / センサー / ロボット制御設備", f_note)
    
    # Sub-container 1B: 現場エッジ & ゲートウェイ
    draw_cell_container(ws1, 19, 2, 24, 9, "現場エッジ・ゲートウェイ層", f_box_head_slate, f_bg_white)
    ws1.write(20, 2, "ローカル暗号化・データ前処理 (Greengrass)", f_note)
    
    # Sub-container 1C: 拠点オフィス・業務端末 (IT系)
    draw_cell_container(ws1, 26, 2, 30, 9, "拠点オフィス・現場端末 (IT系)", f_box_head_slate, f_bg_white)

    # ---------------- Pillar 2: Interconnect Network ----------------
    draw_cell_container(ws1, 5, 11, 31, 16, "【ネットワーク中継層】", f_box_head_dark, f_bg_subnet_pri, head_rows=1)
    
    # Route A: Direct Connect
    draw_cell_container(ws1, 7, 12, 14, 15, "【主回線】専用線接続\nAWS Direct Connect\n(1Gbps/10Gbps帯域)", f_box_head_slate, f_bg_white, head_rows=2)
    
    # Route B: VPN Backup
    draw_cell_container(ws1, 16, 12, 22, 15, "【副回線】暗号化VPN\nSite-to-Site VPN\n(IPsec 自動切替冗長)", f_box_head_slate, f_bg_white, head_rows=2)
    
    # Route C: Cellular / Public
    draw_cell_container(ws1, 24, 12, 30, 15, "【外部網】モバイル回線\n4G/5G・公衆網\n(TLS 1.3 暗号化)", f_box_head_slate, f_bg_white, head_rows=2)

    # ---------------- Pillar 3: AWS Cloud ----------------
    draw_cell_container(ws1, 5, 17, 31, 40, "【AWS クラウド領域】東京リージョン (ap-northeast-1) Multi-AZ クラウド基盤", f_box_head_dark, f_bg_cloud, head_rows=1)
    
    # Tier 1: データ収集・受入口 (Cols 18 to 22, Rows 7 to 24)
    draw_cell_container(ws1, 7, 18, 24, 22, "データ取込・API受付層\n(Ingestion Tier)", f_box_head_slate, f_bg_white, head_rows=2)
    
    # Tier 2: 業務処理・コンテナ層 (Cols 24 to 28, Rows 7 to 24)
    draw_cell_container(ws1, 7, 24, 24, 28, "データ処理・業務AP層\n(Processing Tier)", f_box_head_slate, f_bg_white, head_rows=2)
    
    # Tier 3: 蓄積・分析・DB層 (Cols 30 to 34, Rows 7 to 24)
    draw_cell_container(ws1, 7, 30, 24, 34, "データ永続化・分析層\n(Storage & DB Tier)", f_box_head_slate, f_bg_white, head_rows=2)
    
    # Tier 4: 外部連携・BI層 (Cols 36 to 39, Rows 7 to 24)
    draw_cell_container(ws1, 7, 36, 24, 39, "外部配信・BI層\n(Delivery Tier)", f_box_head_slate, f_bg_white, head_rows=2)
    
    # Connecting Arrows between tiers in AWS
    ws1.merge_range(14, 23, 16, 23, "▶\n▶", f_flow_arrow)
    ws1.merge_range(14, 29, 16, 29, "▶\n▶", f_flow_arrow)
    ws1.merge_range(14, 35, 16, 35, "▶\n▶", f_flow_arrow)
    
    # Tier 5: 横断管理・セキュリティ (Cols 18 to 39, Rows 26 to 30)
    draw_cell_container(ws1, 26, 18, 30, 39, "【共通基盤】統合セキュリティ・運用監視・ID統制 (CloudWatch / GuardDuty / KMS / IAM)", f_box_head_slate, f_bg_white, head_rows=1)

    # ---------------- Insert Representative Icons on Overview ----------------
    # Pillar 1 icons
    ws1.insert_image(9, 3, 'assets_icons/onprem_factory.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(9, 6, 'assets_icons/onprem_plc.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(13, 3, 'assets_icons/onprem_sensor.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(13, 6, 'assets_icons/onprem_robot.png', {'x_scale': 0.75, 'y_scale': 0.75})
    
    ws1.insert_image(20, 3, 'assets_icons/onprem_edge_ipc.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(20, 6, 'assets_icons/onprem_edge_router.png', {'x_scale': 0.75, 'y_scale': 0.75})
    
    ws1.insert_image(27, 3, 'assets_icons/client_pc.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws1.insert_image(27, 6, 'assets_icons/client_tablet.png', {'x_scale': 0.65, 'y_scale': 0.65})

    # Pillar 2 icons
    ws1.insert_image(10, 13, 'assets_icons/aws_dx.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(18, 13, 'assets_icons/aws_vpn.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(26, 13, 'assets_icons/carrier_cellular.png', {'x_scale': 0.75, 'y_scale': 0.75})

    # Pillar 3 icons (AWS)
    # Tier 1 Ingestion
    ws1.insert_image(10, 19, 'assets_icons/aws_iot_core.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(15, 19, 'assets_icons/aws_apigw.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(20, 19, 'assets_icons/aws_tgw.png', {'x_scale': 0.75, 'y_scale': 0.75})

    # Tier 2 Processing
    ws1.insert_image(10, 25, 'assets_icons/aws_kinesis.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(15, 25, 'assets_icons/aws_lambda.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(20, 25, 'assets_icons/aws_ecs.png', {'x_scale': 0.75, 'y_scale': 0.75})

    # Tier 3 Storage
    ws1.insert_image(10, 31, 'assets_icons/aws_s3.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(15, 31, 'assets_icons/aws_rds.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(20, 31, 'assets_icons/aws_dynamodb.png', {'x_scale': 0.75, 'y_scale': 0.75})

    # Tier 4 Delivery & External
    ws1.insert_image(10, 37, 'assets_icons/aws_cloudfront.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(15, 37, 'assets_icons/aws_opensearch.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws1.insert_image(20, 37, 'assets_icons/carrier_internet.png', {'x_scale': 0.75, 'y_scale': 0.75})

    # Tier 5 Security & Governance
    ws1.insert_image(27, 20, 'assets_icons/aws_iam.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws1.insert_image(27, 24, 'assets_icons/aws_kms.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws1.insert_image(27, 28, 'assets_icons/aws_waf.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws1.insert_image(27, 32, 'assets_icons/aws_guardduty.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws1.insert_image(27, 36, 'assets_icons/aws_cloudwatch.png', {'x_scale': 0.65, 'y_scale': 0.65})

    # ---------------- Bottom: System Overview Specification Table ----------------
    row_t1 = 33
    ws1.merge_range(row_t1, 1, row_t1, 40, "システム概要・主要サブシステム仕様一覧", f_sec_title)
    row_t1 += 1
    
    t1_cols = [
        ("No.", 2),
        ("システム区分", 4),
        ("サブシステム名称", 6),
        ("主要機能・処理役割", 11),
        ("採用主要コンポーネント / サービス", 9),
        ("可用性・冗長化方針", 5),
        ("備考・特記事項", 4)
    ]
    cur_c = 1
    for hname, cspan in t1_cols:
        if cspan == 1:
            ws1.write(row_t1, cur_c, hname, f_tbl_head)
        else:
            ws1.merge_range(row_t1, cur_c, row_t1, cur_c + cspan - 1, hname, f_tbl_head)
        cur_c += cspan
    ws1.set_row(row_t1, 22)
    row_t1 += 1
    
    overview_specs = [
        ("01", "現場OT設備", "工場製造ライン・センシング", "製造設備稼働値、温度・振動・電流データの秒周期収集、PLC通信", "三菱電機/オムロン PLC, 各種IoTセンサー, Modbus TCP", "設備側二重化・コールドスタンバイ予備機", "OT系独立NW"),
        ("02", "現場エッジ", "エッジゲートウェイ・前処理", "生データのクレンジング、ローカル一次判定、通信途絶時の一時バッファリング", "産業用エッジPC (Linux), AWS IoT Greengrass v2", "ローカルストレージ (SSD) による72hデータ蓄積", "X.509証明書認証"),
        ("03", "現場IT端末", "現場業務・実績入力クライアント", "製造指示の確認、作業実績登録、ライン異常アラートの現場即時通知閲覧", "現場防塵タブレット, ハンディターミナル, 業務PC", "工場内無線LAN AP 複数台冗長配置", "Web/API経由"),
        ("04", "NW中継", "専用線・閉域ハイブリッド接続", "現場 ⇔ AWS間の高信頼・低遅延データ伝送、大容量テレメトリ転送", "AWS Direct Connect (1Gbps) + Site-to-Site VPN (自動切替)", "BGPによる自動フェイルオーバー (SLA 99.9%)", "完全閉域ルーティング"),
        ("05", "AWS取込", "IoTメッセージング・API受付", "現場数万デバイスからの高並列MQTT接続受付、業務REST API認証", "AWS IoT Core, Amazon API Gateway, AWS WAF", "AWS マネージドMulti-AZ高可用性 (SLA 99.95%)", "TLS1.3 / mTLS"),
        ("06", "AWS処理", "リアルタイムストリーム・業務AP", "時系列データのリアルタイム異常検知、マイクロサービス業務ロジック実行", "Amazon Kinesis Data Streams, AWS Lambda, Amazon ECS", "オートスケーリング (負荷連動自動拡張)", "サーバーレス中心"),
        ("07", "AWS蓄積", "データレイク・基幹DB基盤", "時系列生データの永続化保管、基幹マスタ・トランザクション高速処理", "Amazon S3 (Data Lake), Amazon Aurora PostgreSQL Multi-AZ", "Multi-AZ 自動フェイルオーバー (<30秒), S3 耐久性 99.999999999%", "SSE-KMS暗号化"),
        ("08", "統制監視", "統合運用・セキュリティ監視", "リソース死活・パフォーマンス監視、監査ログ記録、脅威検知・自動通報", "Amazon CloudWatch, AWS CloudTrail, GuardDuty, AWS KMS", "24/365 自動アラート通知 (Slack / メール)", "セキュリティ統制")
    ]
    for r_idx, (no, s_cat, s_name, s_role, s_comp, s_ha, s_note) in enumerate(overview_specs):
        is_even = (r_idx % 2 == 1)
        c_fmt = f_tbl_cell_zebra if is_even else f_tbl_cell
        c_fmt_c = f_tbl_cell_zebra_center if is_even else f_tbl_cell_center
        
        ws1.merge_range(row_t1, 1, row_t1, 2, no, c_fmt_c)
        ws1.merge_range(row_t1, 3, row_t1, 6, s_cat, c_fmt_c)
        ws1.merge_range(row_t1, 7, row_t1, 12, s_name, c_fmt)
        ws1.merge_range(row_t1, 13, row_t1, 23, s_role, c_fmt)
        ws1.merge_range(row_t1, 24, row_t1, 32, s_comp, c_fmt)
        ws1.merge_range(row_t1, 33, row_t1, 37, s_ha, c_fmt)
        ws1.merge_range(row_t1, 38, row_t1, 41, s_note, c_fmt_c)
        ws1.set_row(row_t1, 20)
        row_t1 += 1


    # =========================================================================
    # SHEET 2: 02_構成図_詳細 (Detailed Architecture Diagram Template)
    # =========================================================================
    ws2 = workbook.add_worksheet('02_構成図_詳細')
    ws2.hide_gridlines(0)
    ws2.set_landscape()
    ws2.set_paper(8) # A3
    ws2.fit_to_pages(1, 0)
    ws2.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    
    ws2.set_column('A:A', 2)
    for c in range(1, 42):
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        ws2.set_column(f'{col_letter}:{col_letter}', 3.4)
    for r in range(4, 85):
        ws2.set_row(r, 18)

    write_meta_header(ws2, "【詳細設計】現場システム（OT/IT/DMZ）× AWS Multi-AZ VPC 詳細構成図", "インフラ・NW・開発エンジニア向け 実装設計図")

    # ---------------- Zone 1: On-Premises Factory Network ----------------
    draw_cell_container(ws2, 5, 1, 33, 10, "【オンプレミス工場・現場ネットワーク】CIDR: 192.168.0.0/16", f_box_head_dark, f_bg_onprem, head_rows=1)
    
    # 1A: OT Isolated Zone (Rows 7 to 15)
    draw_cell_container(ws2, 7, 2, 15, 9, "OT系 隔離制御ネットワーク (192.168.10.0/24)", f_box_head_slate, f_bg_white)
    ws2.write(8, 2, "産業設備・計装制御 (外部直接通信不可 / Modbus TCP)", f_note)
    
    # 1B: IT Field Zone (Rows 17 to 24)
    draw_cell_container(ws2, 17, 2, 24, 9, "IT系 現場業務ネットワーク (192.168.20.0/24)", f_box_head_slate, f_bg_white)
    ws2.write(18, 2, "エッジ処理 & 現場業務端末 (Greengrass / Wi-Fi)", f_note)
    
    # 1C: DMZ Security Zone (Rows 26 to 32)
    draw_cell_container(ws2, 26, 2, 32, 9, "現場 DMZ & 境界セキュリティ (192.168.0.0/24)", f_box_head_slate, f_bg_white)
    ws2.write(27, 2, "次世代UTM FW & Customer GW ルータ (BGPピア)", f_note)

    # On-Premise Icons
    ws2.insert_image(9, 3, 'assets_icons/onprem_plc.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws2.insert_image(9, 6, 'assets_icons/onprem_sensor.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws2.insert_image(12, 3, 'assets_icons/onprem_robot.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws2.insert_image(12, 6, 'assets_icons/onprem_hmi.png', {'x_scale': 0.75, 'y_scale': 0.75})

    ws2.insert_image(19, 3, 'assets_icons/onprem_edge_ipc.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws2.insert_image(19, 6, 'assets_icons/client_tablet.png', {'x_scale': 0.75, 'y_scale': 0.75})

    ws2.insert_image(28, 3, 'assets_icons/onprem_utm_fw.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws2.insert_image(28, 6, 'assets_icons/onprem_edge_router.png', {'x_scale': 0.75, 'y_scale': 0.75})

    # ---------------- Zone 2: Interconnect Network ----------------
    draw_cell_container(ws2, 5, 11, 33, 16, "【相互接続・中継層】", f_box_head_dark, f_bg_subnet_pri, head_rows=1)
    
    draw_cell_container(ws2, 7, 12, 13, 15, "Direct Connect\n専用線ポート (1Gbps)\nVLAN 100", f_box_head_slate, f_bg_white, head_rows=2)
    draw_cell_container(ws2, 15, 12, 22, 15, "Transit Gateway (TGW)\nASN: 64512\nCIDR: 10.254.0.0/16", f_box_head_slate, f_bg_white, head_rows=2)
    draw_cell_container(ws2, 24, 12, 32, 15, "Site-to-Site VPN\n(IPsec 冗長トンネル)\nCustomer GW 対向", f_box_head_slate, f_bg_white, head_rows=2)

    ws2.insert_image(9, 13, 'assets_icons/aws_dx.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws2.insert_image(17, 13, 'assets_icons/aws_tgw.png', {'x_scale': 0.75, 'y_scale': 0.75})
    ws2.insert_image(26, 13, 'assets_icons/aws_vpn.png', {'x_scale': 0.75, 'y_scale': 0.75})

    # ---------------- Zone 3: AWS VPC (Multi-AZ) ----------------
    draw_cell_container(ws2, 5, 17, 33, 40, "【AWS VPC】本番環境 VPC: 10.0.0.0/16 (Tokyo Region: ap-northeast-1)", f_box_head_dark, f_bg_vpc, head_rows=1)

    # AZ-1a Boundary (Cols 18 to 28, Rows 7 to 27)
    draw_cell_container(ws2, 7, 18, 27, 28, "Availability Zone 1a (ap-northeast-1a)", f_box_head_slate, f_bg_cloud, head_rows=1)
    
    # AZ-1a Subnets
    draw_cell_container(ws2, 9, 19, 14, 27, "Public Subnet 1a (10.0.1.0/24) | IGW・ALB・NAT GW", f_box_head_gray, f_bg_subnet_pub)
    ws2.insert_image(10, 20, 'assets_icons/aws_alb.png', {'x_scale': 0.7, 'y_scale': 0.7})
    ws2.insert_image(10, 24, 'assets_icons/aws_natgw.png', {'x_scale': 0.7, 'y_scale': 0.7})
    
    draw_cell_container(ws2, 15, 19, 20, 27, "Private App Subnet 1a (10.0.11.0/24) | ECS Fargate & EC2", f_box_head_gray, f_bg_subnet_pri)
    ws2.insert_image(16, 20, 'assets_icons/aws_ecs.png', {'x_scale': 0.7, 'y_scale': 0.7})
    ws2.insert_image(16, 24, 'assets_icons/aws_ec2.png', {'x_scale': 0.7, 'y_scale': 0.7})
    
    draw_cell_container(ws2, 21, 19, 26, 27, "Private DB Subnet 1a (10.0.21.0/24) | Aurora Primary", f_box_head_gray, f_bg_subnet_db)
    ws2.insert_image(22, 22, 'assets_icons/aws_rds.png', {'x_scale': 0.7, 'y_scale': 0.7})

    # AZ-1c Boundary (Cols 30 to 39, Rows 7 to 27)
    draw_cell_container(ws2, 7, 30, 27, 39, "Availability Zone 1c (ap-northeast-1c)", f_box_head_slate, f_bg_cloud, head_rows=1)
    
    # AZ-1c Subnets
    draw_cell_container(ws2, 9, 31, 14, 38, "Public Subnet 1c (10.0.2.0/24) | ALB・NAT GW", f_box_head_gray, f_bg_subnet_pub)
    ws2.insert_image(10, 32, 'assets_icons/aws_alb.png', {'x_scale': 0.7, 'y_scale': 0.7})
    ws2.insert_image(10, 35, 'assets_icons/aws_natgw.png', {'x_scale': 0.7, 'y_scale': 0.7})
    
    draw_cell_container(ws2, 15, 31, 20, 38, "Private App Subnet 1c (10.0.12.0/24) | ECS Fargate", f_box_head_gray, f_bg_subnet_pri)
    ws2.insert_image(16, 32, 'assets_icons/aws_ecs.png', {'x_scale': 0.7, 'y_scale': 0.7})
    ws2.insert_image(16, 35, 'assets_icons/aws_fargate.png', {'x_scale': 0.7, 'y_scale': 0.7})
    
    draw_cell_container(ws2, 21, 31, 26, 38, "Private DB Subnet 1c (10.0.22.0/24) | Aurora Reader", f_box_head_gray, f_bg_subnet_db)
    ws2.insert_image(22, 33, 'assets_icons/aws_rds.png', {'x_scale': 0.7, 'y_scale': 0.7})

    # Storage & Serverless Tier (Rows 28 to 32, Cols 18 to 39)
    draw_cell_container(ws2, 28, 18, 32, 39, "【マネージド・サーバーレス層】S3 Data Lake / IoT Core / Kinesis / DynamoDB / Lambda", f_box_head_slate, f_bg_white, head_rows=1)
    ws2.insert_image(29, 19, 'assets_icons/aws_iot_core.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws2.insert_image(29, 23, 'assets_icons/aws_kinesis.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws2.insert_image(29, 27, 'assets_icons/aws_lambda.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws2.insert_image(29, 31, 'assets_icons/aws_s3.png', {'x_scale': 0.65, 'y_scale': 0.65})
    ws2.insert_image(29, 35, 'assets_icons/aws_dynamodb.png', {'x_scale': 0.65, 'y_scale': 0.65})

    # ---------------- Bottom: Detailed Component Specifications Table ----------------
    row_d = 35
    ws2.merge_range(row_d, 1, row_d, 40, "詳細設計リソーススペック・IP・ポート・セキュリティ一覧表", f_sec_title)
    row_d += 1
    
    d_cols = [
        ("No.", 2),
        ("リソース名称", 5),
        ("リソース種別", 4),
        ("所属ゾーン / サブネット", 6),
        ("IPアドレス / CIDR", 5),
        ("ポート / プロトコル", 4),
        ("セキュリティグループ / 認証", 6),
        ("冗長化方式 / バックアップ", 4),
        ("役割・設計パラメータ", 5)
    ]
    cur_c = 1
    for hname, cspan in d_cols:
        if cspan == 1:
            ws2.write(row_d, cur_c, hname, f_tbl_head)
        else:
            ws2.merge_range(row_d, cur_c, row_d, cur_c + cspan - 1, hname, f_tbl_head)
        cur_c += cspan
    ws2.set_row(row_d, 22)
    row_d += 1
    
    detail_specs = [
        ("01", "現場 PLC-01", "産業制御器", "工場 OT系ネットワーク", "192.168.10.11/24", "502 / Modbus TCP", "物理スイッチPort閉塞", "コールドスタンバイ予備", "製造ラインA自動制御"),
        ("02", "現場 PLC-02", "産業制御器", "工場 OT系ネットワーク", "192.168.10.12/24", "502 / Modbus TCP", "物理スイッチPort閉塞", "コールドスタンバイ予備", "製造ラインB自動制御"),
        ("03", "現場 Edge IPC", "産業用PC (Linux)", "工場 IT系ネットワーク", "192.168.20.15/24", "8883 / MQTT, 502 / TCP", "iptables (内部発のみ)", "SSDローカルバッファ", "Greengrass v2 / 前処理"),
        ("04", "現場 UTM / FW", "次世代ファイアウォール", "現場 DMZゾーン", "192.168.0.1/24", "All / 厳格ステートフル", "IP/MACバインド", "アクティブ/スタンバイHA", "OT/IT間境界防御"),
        ("05", "拠点 ルータ (CGW)", "Cisco ISR 4331", "現場 DMZゾーン", "192.168.0.254/24", "179/BGP, 500/4500 IPsec", "境界ACL (AWS対向のみ)", "VRRP デュアルルータ", "DX & VPN BGP終端"),
        ("06", "Direct Connect", "物理専用線ポート", "東京コロケーション", "VLAN 100", "802.1Q タグVLAN", "キャリア構内相互接続", "1Gbps 専用線帯域", "オンプレ⇔AWS間主回線"),
        ("07", "Transit Gateway", "AWS TGW", "AWS NW中継層", "10.254.0.0/16", "BGP ルーティング", "TGW Route Table", "AWS Multi-AZ高可用性", "VPC & DX/VPN 統合集約"),
        ("08", "ALB (外部向)", "Application LB", "Public 1a / 1c", "10.0.1.x, 10.0.2.x", "443 / HTTPS", "sg-alb-external", "クロスゾーン負荷分散", "現場端末/外部向けAPI受付"),
        ("09", "NAT Gateway 1a/1c", "NAT ゲートウェイ", "Public 1a / 1c", "10.0.1.x, 10.0.2.x", "送信元NAT変換", "EIP 割当", "AZ個別冗長化 (2台配置)", "プライベートサブネット外向き通信"),
        ("10", "ECS Fargate (AP)", "コンテナタスク", "Private App 1a / 1c", "10.0.11.50, 10.0.12.50", "8080 / HTTP", "sg-ecs-app (ALBからのみ)", "Auto Scaling (2〜10タスク)", "FastAPI / マイクロサービス"),
        ("11", "Aurora PostgreSQL", "Amazon Aurora DB", "Private DB 1a / 1c", "10.0.21.100, 10.0.22.100", "5432 / PostgreSQL", "sg-aurora-db (ECSからのみ)", "Multi-AZ 自動フェイルオーバー", "db.r6g.xlarge, 日次Snapshot"),
        ("12", "AWS IoT Core", "IoTメッセージブローカー", "リージョン共通", "-", "8883 / MQTT over TLS", "X.509 デバイス証明書", "完全マネージド分散基盤", "秒間10,000件テレメトリ受信"),
        ("13", "Amazon S3 Data Lake", "オブジェクトストレージ", "リージョン共通", "-", "443 / HTTPS (VPC Endpoint)", "バケットポリシー & IAM", "耐久性 99.999999999%", "SSE-KMS暗号化, ライフサイクル"),
        ("14", "AWS KMS", "鍵管理サービス", "リージョン共通", "-", "443 / HTTPS", "KMSキーポリシー", "AWS HSM 管理", "データレイク・DB一元暗号鍵"),
        ("15", "CloudWatch Logs", "統合監視・ログ基盤", "リージョン共通", "-", "443 / HTTPS (VPC Endpoint)", "IAM ロール認証", "Multi-AZ 保管", "アラーム検知時 SNS/Slack通知")
    ]
    for r_idx, r_vals in enumerate(detail_specs):
        is_even = (r_idx % 2 == 1)
        c_fmt = f_tbl_cell_zebra if is_even else f_tbl_cell
        c_fmt_c = f_tbl_cell_zebra_center if is_even else f_tbl_cell_center
        
        ws2.merge_range(row_d, 1, row_d, 2, r_vals[0], c_fmt_c)
        ws2.merge_range(row_d, 3, row_d, 7, r_vals[1], c_fmt)
        ws2.merge_range(row_d, 8, row_d, 11, r_vals[2], c_fmt_c)
        ws2.merge_range(row_d, 12, row_d, 17, r_vals[3], c_fmt)
        ws2.merge_range(row_d, 18, row_d, 22, r_vals[4], f_tbl_cell_code)
        ws2.merge_range(row_d, 23, row_d, 26, r_vals[5], f_tbl_cell_code)
        ws2.merge_range(row_d, 27, row_d, 32, r_vals[6], c_fmt)
        ws2.merge_range(row_d, 33, row_d, 36, r_vals[7], c_fmt)
        ws2.merge_range(row_d, 37, row_d, 41, r_vals[8], c_fmt)
        ws2.set_row(row_d, 20)
        row_d += 1


    # =========================================================================
    # SHEET 3: 03_パーツ集_概要用 (Overview Parts & Palette)
    # =========================================================================
    ws3 = workbook.add_worksheet('03_パーツ集_概要用')
    ws3.hide_gridlines(0)
    ws3.set_landscape()
    ws3.set_paper(8)
    ws3.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    
    ws3.set_column('A:A', 2)
    for c in range(1, 42):
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        ws3.set_column(f'{col_letter}:{col_letter}', 3.4)
    for r in range(4, 80):
        ws3.set_row(r, 18)

    write_meta_header(ws3, "【貼り付け用パーツ集】概要構成図用 マクロ枠・階層カード・大口径線種・凡例", "概要図作成用 コピペパレット")

    # Section A: マクロ境界コンテナ
    row_p3 = 5
    ws3.merge_range(row_p3, 1, row_p3, 40, "A. 概要図用 マクロ境界コンテナ (選択してコピー＆ペーストして使用)", f_sec_title)
    row_p3 += 1
    
    draw_cell_container(ws3, row_p3, 1, row_p3 + 8, 12, "【マクロ枠】現場・スマート工場領域 (OT/IT)", f_box_head_dark, f_bg_onprem)
    draw_cell_container(ws3, row_p3, 14, row_p3 + 8, 25, "【マクロ枠】通信キャリア専用線・閉域網 (Direct Connect)", f_box_head_slate, f_bg_subnet_pri)
    draw_cell_container(ws3, row_p3, 27, row_p3 + 8, 40, "【マクロ枠】AWS Cloud (東京リージョン ap-northeast-1)", f_box_head_dark, f_bg_cloud)
    
    ws3.write(row_p3 + 2, 2, "← 工場・プラント・拠点を\n　包括する外枠として利用", f_note)
    ws3.write(row_p3 + 2, 15, "← 専用線・VPN・閉域網の\n　中継領域外枠として利用", f_note)
    ws3.write(row_p3 + 2, 28, "← クラウド環境全体を\n　包括する外枠として利用", f_note)
    row_p3 += 10
    
    # Section B: マクロ階層カード (Macro Tier Badges)
    ws3.merge_range(row_p3, 1, row_p3, 40, "B. 概要図用 主要機能階層カード (ドラッグまたはコピーして概要図に配置)", f_sec_title)
    row_p3 += 1
    
    macro_cards = [
        ('macro_factory.png', 1),
        ('macro_office.png', 9),
        ('macro_dx.png', 17),
        ('macro_internet.png', 25),
        ('macro_aws_region.png', 33),
    ]
    for img_name, col_pos in macro_cards:
        ws3.insert_image(row_p3, col_pos, f'assets_icons/{img_name}', {'x_scale': 0.85, 'y_scale': 0.85})
    row_p3 += 4

    macro_cards_2 = [
        ('macro_tier_ingestion.png', 1),
        ('macro_tier_compute.png', 9),
        ('macro_tier_storage.png', 17),
        ('macro_tier_security.png', 25),
        ('macro_external_saas.png', 33),
    ]
    for img_name, col_pos in macro_cards_2:
        ws3.insert_image(row_p3, col_pos, f'assets_icons/{img_name}', {'x_scale': 0.85, 'y_scale': 0.85})
    row_p3 += 5

    # Section C: 概要用 代表AWS & 現場アイコン
    ws3.merge_range(row_p3, 1, row_p3, 40, "C. 概要図用 代表サービス & 設備アイコン (概要図にそのままコピー可能)", f_sec_title)
    row_p3 += 1
    
    overview_icons_sample = [
        ('onprem_factory.png', '工場プラント', 1),
        ('onprem_plc.png', '現場制御PLC', 5),
        ('onprem_edge_ipc.png', 'エッジPC', 9),
        ('aws_dx.png', 'Direct Connect', 13),
        ('aws_tgw.png', 'Transit GW', 17),
        ('aws_iot_core.png', 'IoT Core', 21),
        ('aws_ecs.png', 'コンテナECS', 25),
        ('aws_lambda.png', 'Lambda関数', 29),
        ('aws_s3.png', 'S3データレイク', 33),
        ('aws_rds.png', 'Aurora DB', 37),
    ]
    for img_file, label, c_idx in overview_icons_sample:
        ws3.insert_image(row_p3, c_idx, f'assets_icons/{img_file}', {'x_scale': 0.75, 'y_scale': 0.75})
        ws3.merge_range(row_p3 + 5, c_idx, row_p3 + 5, c_idx + 2, label, f_tbl_cell_center)
    row_p3 += 7

    # Section D: 概要用 大口径接続線 & データフロー
    ws3.merge_range(row_p3, 1, row_p3, 40, "D. 概要図用 データ流通・中継線種 (太線・双方向・専用線)", f_sec_title)
    row_p3 += 1
    
    ws3.insert_image(row_p3, 1, 'assets_icons/line_direct_connect.png', {'x_scale': 0.9, 'y_scale': 0.9})
    ws3.insert_image(row_p3, 15, 'assets_icons/line_solid_sync.png', {'x_scale': 0.9, 'y_scale': 0.9})
    ws3.insert_image(row_p3, 28, 'assets_icons/line_dashed_async.png', {'x_scale': 0.9, 'y_scale': 0.9})
    row_p3 += 3
    
    ws3.insert_image(row_p3, 1, 'assets_icons/line_vpn_tunnel.png', {'x_scale': 0.9, 'y_scale': 0.9})
    ws3.insert_image(row_p3, 15, 'assets_icons/line_bidirectional.png', {'x_scale': 0.9, 'y_scale': 0.9})
    ws3.insert_image(row_p3, 28, 'assets_icons/line_dotted_mgmt.png', {'x_scale': 0.9, 'y_scale': 0.9})
    row_p3 += 4

    # Section E: エグゼクティブ凡例ブロック
    ws3.merge_range(row_p3, 1, row_p3, 40, "E. 概要図用 エグゼクティブ凡例ブロック (完成図の右下や空き領域に配置)", f_sec_title)
    row_p3 += 1
    
    draw_cell_container(ws3, row_p3, 1, row_p3 + 6, 20, "【概要図 凡例】通信・ネットワーク種別", f_box_head_slate, f_bg_white)
    ws3.write(row_p3 + 1, 2, "━━━ [太実線] 専用線接続 (AWS Direct Connect 1Gbps/10Gbps)", f_note)
    ws3.write(row_p3 + 2, 2, "━━▶ [実線矢印] 同期通信 (Web / REST API / HTTPS 暗号化)", f_note)
    ws3.write(row_p3 + 3, 2, "----▶ [破線矢印] 非同期通信 (IoT テレメトリ / MQTT / キュー)", f_note)
    ws3.write(row_p3 + 4, 2, "◀━━▶ [双方向] 双方向制御・同期 (Modbus TCP / 設備制御)", f_note)
    ws3.write(row_p3 + 5, 2, "・・・・▶ [点線] 運用監視・メトリクス・ログ収集 (CloudWatch)", f_note)

    draw_cell_container(ws3, row_p3, 22, row_p3 + 6, 40, "【概要図 凡例】環境・セキュリティ境界", f_box_head_slate, f_bg_white)
    ws3.write(row_p3 + 1, 23, "■ 現場OT領域 : 外部インターネットから完全物理/論理隔離", f_note)
    ws3.write(row_p3 + 2, 23, "■ 現場IT領域 : 拠点内閉域ネットワーク (認証端末のみ接続許可)", f_note)
    ws3.write(row_p3 + 3, 23, "■ AWS VPC領域 : 仮想プライベート網 (インターネット非公開・閉域ルーティング)", f_note)
    ws3.write(row_p3 + 4, 23, "■ 全通信暗号化 : TLS 1.3 / IPsec / SSE-KMS 保存時暗号化を標準適用", f_note)
    ws3.write(row_p3 + 5, 23, "■ 冗長化方針 : Multi-AZ 構成により単一障害点 (SPOF) を完全排除", f_note)


    # =========================================================================
    # SHEET 4: 04_パーツ集_詳細用 (Detailed Parts & Asset Catalog)
    # =========================================================================
    ws4 = workbook.add_worksheet('04_パーツ集_詳細用')
    ws4.hide_gridlines(0)
    ws4.set_landscape()
    ws4.set_paper(8)
    ws4.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    
    ws4.set_column('A:A', 2)
    for c in range(1, 42):
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        ws4.set_column(f'{col_letter}:{col_letter}', 3.4)
    for r in range(4, 130):
        ws4.set_row(r, 18)

    write_meta_header(ws4, "【貼り付け用パーツ集】詳細構成図用 アイコン・サブネット枠・線種・ポートタグ・表", "精密設計図作成用 全パーツカタログ")

    row_p4 = 5
    # Section A: 詳細ネットワークコンテナ枠
    ws4.merge_range(row_p4, 1, row_p4, 40, "A. 詳細設計用 ネットワーク境界コンテナ枠 (コピーしてサイズ調整して利用)", f_sec_title)
    row_p4 += 1
    
    draw_cell_container(ws4, row_p4, 1, row_p4 + 7, 10, "【枠】VPC 10.0.0.0/16", f_box_head_dark, f_bg_vpc)
    draw_cell_container(ws4, row_p4, 12, row_p4 + 7, 21, "【枠】Public Subnet (10.0.1.0/24)", f_box_head_gray, f_bg_subnet_pub)
    draw_cell_container(ws4, row_p4, 23, row_p4 + 7, 32, "【枠】Private App Subnet (10.0.11.0/24)", f_box_head_slate, f_bg_subnet_pri)
    draw_cell_container(ws4, row_p4, 34, row_p4 + 7, 41, "【枠】Private DB Subnet", f_box_head_dark, f_bg_subnet_db)
    row_p4 += 9

    # Section B: AWS サービス全アイコンカタログ (40+ items)
    ws4.merge_range(row_p4, 1, row_p4, 40, "B. AWS サービス詳細アイコンカタログ (すべてのアイコンをCtrl+Cで即座にコピー可能)", f_sec_title)
    row_p4 += 1
    
    aws_categories = [
        ("【AWS Compute & Containers】", [
            ('aws_ec2.png', 'EC2\n仮想サーバ'),
            ('aws_ecs.png', 'ECS\nコンテナ'),
            ('aws_eks.png', 'EKS\nK8s基盤'),
            ('aws_lambda.png', 'Lambda\n関数実行'),
            ('aws_fargate.png', 'Fargate\nサーバレス'),
            ('aws_batch.png', 'Batch\nバッチ処理'),
        ]),
        ("【AWS Storage & Content Delivery】", [
            ('aws_s3.png', 'S3\nオブジェクト'),
            ('aws_glacier.png', 'Glacier\n長期保管'),
            ('aws_efs.png', 'EFS\n共有ストレージ'),
            ('aws_ebs.png', 'EBS\nブロック'),
            ('aws_cloudfront.png', 'CloudFront\nCDN配信'),
        ]),
        ("【AWS Database & Analytics】", [
            ('aws_rds.png', 'RDS/Aurora\nリレーショナル'),
            ('aws_dynamodb.png', 'DynamoDB\nNoSQL高速'),
            ('aws_elasticache.png', 'ElastiCache\nキャッシュ'),
            ('aws_opensearch.png', 'OpenSearch\n検索・ログ'),
            ('aws_documentdb.png', 'DocDB\nドキュメント'),
        ]),
        ("【AWS Networking & Gateway】", [
            ('aws_vpc.png', 'VPC\n仮想ネットワーク'),
            ('aws_alb.png', 'ALB\nアプリ負荷分散'),
            ('aws_nlb.png', 'NLB\nNW負荷分散'),
            ('aws_route53.png', 'Route 53\nクラウドDNS'),
            ('aws_tgw.png', 'Transit GW\nNW統合集約'),
            ('aws_dx.png', 'Direct Connect\n専用線接続'),
            ('aws_dxgw.png', 'DX GW\nDXゲートウェイ'),
            ('aws_vpn.png', 'Site VPN\nIPsec VPN'),
            ('aws_natgw.png', 'NAT GW\n送信元NAT'),
            ('aws_igw.png', 'Internet GW\nインターネット出入口'),
            ('aws_apigw.png', 'API Gateway\nREST/HTTP'),
            ('aws_privatelink.png', 'PrivateLink\nVPC Endpoint'),
        ]),
        ("【AWS IoT & Streaming & Messaging】", [
            ('aws_iot_core.png', 'IoT Core\n接続ブローカー'),
            ('aws_greengrass.png', 'Greengrass\nエッジ実行基盤'),
            ('aws_sitewise.png', 'SiteWise\n現場設備収集'),
            ('aws_kinesis.png', 'Kinesis\nストリーム'),
            ('aws_sqs.png', 'SQS\nキュー'),
            ('aws_sns.png', 'SNS\n通知配信'),
            ('aws_eventbridge.png', 'EventBridge\nイベントバス'),
            ('aws_stepfunctions.png', 'StepFunctions\nワークフロー'),
        ]),
        ("【AWS Security, Identity & Observability】", [
            ('aws_iam.png', 'IAM\n権限・認証'),
            ('aws_kms.png', 'KMS\n暗号鍵管理'),
            ('aws_secrets.png', 'Secrets Mgr\n認証情報保護'),
            ('aws_waf.png', 'AWS WAF\nWeb防御'),
            ('aws_guardduty.png', 'GuardDuty\n脅威検知'),
            ('aws_cloudwatch.png', 'CloudWatch\nメトリクス監視'),
            ('aws_cloudtrail.png', 'CloudTrail\n監査ログ'),
            ('aws_ssm.png', 'SSM\n運用リモート管理'),
        ])
    ]

    for cat_title, icon_list in aws_categories:
        ws4.merge_range(row_p4, 1, row_p4, 40, cat_title, f_sec_sub)
        row_p4 += 1
        chunk_size = 6
        for chunk_idx in range(0, len(icon_list), chunk_size):
            chunk = icon_list[chunk_idx:chunk_idx + chunk_size]
            for idx, (img_file, label) in enumerate(chunk):
                col_start = 1 + idx * 6
                ws4.insert_image(row_p4, col_start + 1, f'assets_icons/{img_file}', {'x_scale': 0.75, 'y_scale': 0.75})
                ws4.merge_range(row_p4 + 5, col_start, row_p4 + 6, col_start + 5, label, f_tbl_cell_center)
            row_p4 += 8

    # Section C: 現場・OT・IT機器全アイコンカタログ
    ws4.merge_range(row_p4, 1, row_p4, 40, "C. 現場・工場OT・拠点IT機器アイコンカタログ (コピーして詳細図に配置)", f_sec_title)
    row_p4 += 1
    
    onprem_categories = [
        ("【現場・施設 & OT産業制御設備】", [
            ('onprem_factory.png', 'Factory\n工場プラント'),
            ('onprem_office.png', 'Office\n本社・支社'),
            ('onprem_localdc.png', 'Local DC\n自社データセンタ'),
            ('onprem_plc.png', 'PLC\n産業コントローラ'),
            ('onprem_sensor.png', 'Sensor\n計装・IoTセンサー'),
            ('onprem_robot.png', 'Robot\n産業用ロボット'),
        ]),
        ("【現場エッジ・通信ITインフラ & 端末】", [
            ('onprem_edge_ipc.png', 'Edge IPC\n産業用エッジPC'),
            ('onprem_iot_gw.png', 'IoT GW\n現場ゲートウェイ'),
            ('onprem_local_server.png', 'Local Server\n物理サーバ'),
            ('onprem_utm_fw.png', 'UTM / FW\n境界セキュリティ'),
            ('onprem_edge_router.png', 'Edge Router\n拠点ルータ'),
            ('onprem_l3_switch.png', 'L3 Switch\n集約スイッチ'),
        ]),
        ("【現場クライアント・作業端末 & 通信網】", [
            ('client_tablet.png', 'Tablet\n防塵タブレット'),
            ('client_handy.png', 'Handy\nハンディ端末'),
            ('client_pc.png', 'Client PC\n管理者PC'),
            ('carrier_directlink.png', 'Direct Link\n専用線キャリア'),
            ('carrier_closedvpn.png', 'Closed VPN\nIP-VPN閉域網'),
            ('carrier_cellular.png', '5G / LTE\nモバイル回線'),
        ])
    ]
    for cat_title, icon_list in onprem_categories:
        ws4.merge_range(row_p4, 1, row_p4, 40, cat_title, f_sec_sub)
        row_p4 += 1
        for idx, (img_file, label) in enumerate(icon_list):
            col_start = 1 + idx * 6
            ws4.insert_image(row_p4, col_start + 1, f'assets_icons/{img_file}', {'x_scale': 0.75, 'y_scale': 0.75})
            ws4.merge_range(row_p4 + 5, col_start, row_p4 + 6, col_start + 5, label, f_tbl_cell_center)
        row_p4 += 8

    # Section D: プロトコル・ポートタグ & ステータスバッジ
    ws4.merge_range(row_p4, 1, row_p4, 40, "D. 通信プロトコル・ポートタグ & ステータスバッジ (通信線上に重ねて配置)", f_sec_title)
    row_p4 += 1
    
    ptags_1 = [
        ('tag_https_443.png', 1),
        ('tag_mqtt_8883.png', 7),
        ('tag_modbus_502.png', 13),
        ('tag_postgres_5432.png', 19),
        ('tag_mysql_3306.png', 25),
        ('tag_ssh_22.png', 31),
        ('tag_dx_1g.png', 37),
    ]
    for img_file, c_pos in ptags_1:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.9, 'y_scale': 0.9})
    row_p4 += 2

    ptags_2 = [
        ('tag_opcua_4840.png', 1),
        ('tag_rest_api.png', 7),
        ('tag_dns_53.png', 13),
        ('tag_syslog_514.png', 19),
        ('tag_ntp_123.png', 25),
        ('tag_rdp_3389.png', 31),
        ('tag_ipsec_vpn.png', 37),
    ]
    for img_file, c_pos in ptags_2:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.9, 'y_scale': 0.9})
    row_p4 += 3

    # Status badges
    ws4.write(row_p4, 1, "【ステータス・ゾーンバッジ】", f_sec_sub)
    row_p4 += 1
    
    sbadges_1 = [
        ('badge_prod.png', 1),
        ('badge_stg.png', 7),
        ('badge_dev.png', 13),
        ('badge_multiaz.png', 19),
        ('badge_active.png', 25),
        ('badge_standby.png', 31),
        ('badge_encrypted.png', 37),
    ]
    for img_file, c_pos in sbadges_1:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.9, 'y_scale': 0.9})
    row_p4 += 2

    sbadges_2 = [
        ('badge_zone_ot.png', 1),
        ('badge_zone_it.png', 7),
        ('badge_zone_dmz.png', 13),
        ('badge_zone_public.png', 19),
        ('badge_zone_private.png', 25),
        ('badge_zone_db.png', 31),
        ('badge_managed.png', 37),
    ]
    for img_file, c_pos in sbadges_2:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.9, 'y_scale': 0.9})
    row_p4 += 4

    # Section E: 詳細線種ストリップ
    ws4.merge_range(row_p4, 1, row_p4, 40, "E. 詳細線種ストリップ集 (同期・非同期・専用線・VPN・監視・双方向)", f_sec_title)
    row_p4 += 1
    
    line_strips = [
        ('line_solid_sync.png', 1),
        ('line_dashed_async.png', 21),
    ]
    for img_file, c_pos in line_strips:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.95, 'y_scale': 0.95})
    row_p4 += 3

    line_strips_2 = [
        ('line_direct_connect.png', 1),
        ('line_vpn_tunnel.png', 21),
    ]
    for img_file, c_pos in line_strips_2:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.95, 'y_scale': 0.95})
    row_p4 += 3

    line_strips_3 = [
        ('line_bidirectional.png', 1),
        ('line_dotted_mgmt.png', 21),
    ]
    for img_file, c_pos in line_strips_3:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.95, 'y_scale': 0.95})
    row_p4 += 3

    line_strips_4 = [
        ('line_dash_dot_db.png', 1),
    ]
    for img_file, c_pos in line_strips_4:
        ws4.insert_image(row_p4, c_pos, f'assets_icons/{img_file}', {'x_scale': 0.95, 'y_scale': 0.95})
    row_p4 += 4

    # Section F: 詳細リソース仕様表の空テンプレート（コピペ用）
    ws4.merge_range(row_p4, 1, row_p4, 40, "F. 詳細リソース設計一覧表 空テンプレート (コピーして新規システム表として利用)", f_sec_title)
    row_p4 += 1
    
    cur_c = 1
    for hname, cspan in d_cols:
        if cspan == 1:
            ws4.write(row_p4, cur_c, hname, f_tbl_head)
        else:
            ws4.merge_range(row_p4, cur_c, row_p4, cur_c + cspan - 1, hname, f_tbl_head)
        cur_c += cspan
    ws4.set_row(row_p4, 22)
    row_p4 += 1
    
    # 5 blank rows
    for r_idx in range(5):
        is_even = (r_idx % 2 == 1)
        c_fmt = f_tbl_cell_zebra if is_even else f_tbl_cell
        c_fmt_c = f_tbl_cell_zebra_center if is_even else f_tbl_cell_center
        
        ws4.merge_range(row_p4, 1, row_p4, 2, f"{r_idx+1:02d}", c_fmt_c)
        ws4.merge_range(row_p4, 3, row_p4, 7, "", c_fmt)
        ws4.merge_range(row_p4, 8, row_p4, 11, "", c_fmt_c)
        ws4.merge_range(row_p4, 12, row_p4, 17, "", c_fmt)
        ws4.merge_range(row_p4, 18, row_p4, 22, "", f_tbl_cell_code)
        ws4.merge_range(row_p4, 23, row_p4, 26, "", f_tbl_cell_code)
        ws4.merge_range(row_p4, 27, row_p4, 32, "", c_fmt)
        ws4.merge_range(row_p4, 33, row_p4, 36, "", c_fmt)
        ws4.merge_range(row_p4, 37, row_p4, 41, "", c_fmt)
        ws4.set_row(row_p4, 20)
        row_p4 += 1

    # Close workbook
    workbook.close()
    
    # Also create an English alias copy for CLI convenience
    shutil.copyfile(filename, "AWS_Hybrid_Architecture_Template.xlsx")
    print(f"Template workbook successfully generated: {filename}")
    print("English alias copy created: AWS_Hybrid_Architecture_Template.xlsx")

if __name__ == '__main__':
    create_system_architecture_template()
