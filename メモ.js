/**
 * データベース構造 (Database.xlsx / Google Sheets)
 * 
 * ■ 職員情報 (SHEETS.STAFF)
 * A: id
 * B: name (氏名)
 * C: role (基本職名)
 * D: display_order (表示順)
 * E: email (メールアドレス)
 * F: grade (学年)
 * G: dept1 (分掌1)
 * H: dept2 (分掌2)
 * I: dept3 (分掌3)
 * J: subject (教科)
 * K: role_type (役割: 担任, 副担任, 学年所属, 管理職)
 * L: chief_type (主任: 学年主任, 学年副主任, 分掌主任, 分掌副主任)
 * 
 * ■ 行事 (SHEETS.DAILY)
 * A: id
 * B: date (日付)
 * C: time (時間)
 * D: content (内容)
 * E: note (備考)
 * 
 * ■ 出張等 (SHEETS.TRIP)
 * A: id
 * B: date (日付)
 * C: staff_name (氏名)
 * D: purpose (用務)
 * E: location (行先)
 * F: time (期間)
 * G: note (備考)
 * 
 * ■ 休暇等 (SHEETS.LEAVE)
 * A: id
 * B: date (日付)
 * C: staff_name (氏名)
 * D: type (休暇種別)
 * E: time (期間)
 * F: note (備考)
 * 
 * ■ 会議 (SHEETS.MEETING)
 * A: id
 * B: date (日付)
 * C: meeting_name (会議名)
 * D: time_period (時間帯)
 * E: place (場所)
 * 
 * ■ 伝達事項 (SHEETS.ANNOUNCE)
 * A: id
 * B: date (日付)
 * C: type (種別)
 * D: priority (重要度)
 * E: target (対象)
 * F: content (内容)
 * G: reporter (連絡者)
 * 
 * ■ 定例会議 (SHEETS.FIXED_MEETING)
 * A: day_of_week (曜日)
 * B: meeting_name (会議名)
 * C: time (時間)
 * D: place (場所)
 * 
 * ■ 特別教室予約 (SHEETS.ROOM)
 * A: id
 * B: date (日付)
 * C: period (時限)
 * D: room_name (教室名)
 * E: content (内容)
 * F: reserver (予約者)
 * 
 * ■ 特別教室固定 (SHEETS.FIXED_CLASS)
 * A: day_of_week (曜日)
 * B: period (時限)
 * C: room_name (教室名)
 * D: content (内容)
 * E: reserver (使用教員)
 * 
 * ■ タスク (SHEETS.TASK)
 * A: id
 * B: name (担当者)
 * C: role (ロール: 未使用)
 * D: content (内容)
 * E: due_date (期限)
 * F: check (ステータス: 0=未着手, 1=着手中, 2=完了)
 */