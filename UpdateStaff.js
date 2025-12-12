function updateStaffListOneTime() {
    const staffData = [
        { id: '01', name: '船津　勇一', order: 1 }, { id: '02', name: '長石　庄一郎', order: 2 }, { id: '03', name: '佐藤　博義', order: 3 }, { id: '04', name: '裏　久代', order: 4 }, { id: '05', name: '阿部　映真', order: 5 },
        { id: '06', name: '柚木　達也', order: 6 }, { id: '07', name: '髙倉　圭一', order: 7 }, { id: '08', name: '安元　正彦', order: 8 }, { id: '09', name: '宮原　徹', order: 9 }, { id: '10', name: '堀田　秀俊', order: 10 },
        { id: '11', name: '江藤　　賢', order: 11 }, { id: '12', name: '金谷　昭二', order: 12 }, { id: '13', name: '橋内　和彦', order: 13 }, { id: '14', name: '堀谷　桂', order: 14 }, { id: '15', name: '森　佐和美', order: 15 },
        { id: '16', name: '工藤　圭介', order: 16 }, { id: '17', name: '永松　寛明', order: 17 }, { id: '18', name: '小川　尚志', order: 18 }, { id: '19', name: '田北　俊郎', order: 19 }, { id: '20', name: '堀　竜大', order: 20 },
        { id: '21', name: '高田　裕介', order: 21 }, { id: '22', name: '水江　友和', order: 22 }, { id: '23', name: '衛藤　加奈', order: 23 }, { id: '24', name: '西山　幹子', order: 24 }, { id: '25', name: '藤野　真也', order: 25 },
        { id: '26', name: '久保　修平', order: 26 }, { id: '27', name: '小笠原　陽華', order: 27 }, { id: '28', name: '工藤　督右', order: 28 }, { id: '29', name: '藤丸　拓也', order: 29 }, { id: '30', name: '秦　ひとみ', order: 30 },
        { id: '31', name: '三浦　裕希', order: 31 }, { id: '32', name: '大友　宗一郎', order: 32 }, { id: '33', name: '秋吉　豊', order: 33 }, { id: '34', name: '安倍　あいみ', order: 34 }, { id: '35', name: '伊藤　大貴', order: 35 },
        { id: '36', name: '田島　幸太郎', order: 36 }, { id: '37', name: '古田　義貴', order: 37 }, { id: '38', name: '佐藤　陽大', order: 38 }, { id: '39', name: '小池　佑輔', order: 39 }, { id: '40', name: '樽本　有貴', order: 40 },
        { id: '41', name: '中野　豊', order: 41 }, { id: '42', name: '長野　真結', order: 42 }, { id: '43', name: '一井　ひろえ', order: 43 }, { id: '44', name: '玉ノ井　遥廉', order: 44 }, { id: '45', name: '芳友　麻理子', order: 45 },
        { id: '46', name: '佐藤　利佳', order: 46 }, { id: '47', name: '信岡　大喜', order: 47 }, { id: '48', name: '川﨑　みゆき', order: 48 }, { id: '49', name: '中島　貴之', order: 49 }, { id: '50', name: '奥野　沙耶', order: 50 },
        { id: '51', name: '千葉　優希', order: 51 }, { id: '52', name: '岩永　明日翔', order: 52 }, { id: '53', name: '梅木　美嶺', order: 53 }, { id: '54', name: '志賀　真輝', order: 54 }, { id: '55', name: '首藤　鼓太郎', order: 55 },
        { id: '56', name: '野口　大', order: 56 }, { id: '57', name: '宮﨑　恵子', order: 57 }, { id: '58', name: '長木　佳子', order: 58 }, { id: '59', name: '馬場　舜夏', order: 59 }, { id: '60', name: '轟　睦子', order: 60 },
        { id: '61', name: '梶原　希央', order: 61 }, { id: '62', name: '大曲　直子', order: 62 }, { id: '63', name: 'レミー・ファーニス', order: 63 }, { id: '64', name: '小溝　晴美', order: 64 }, { id: '65', name: '宮﨑　晶子', order: 65 },
        { id: '66', name: '羽野　紀美子', order: 66 }
    ];

    const ss = getSS();
    const sheet = ss.getSheetByName(SHEETS.STAFF);
    if (!sheet) throw new Error('STAFF Sheet not found');

    // Clear old data (Except Header)
    // Assuming Header is Row 1.
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 12).clearContent();
    }

    // Prepare Rows
    // 0:id, 1:name, 2:role, 3:order, 4:email, 5:grade, 6:dept1, 7:dept2, 8:dept3, 9:subject, 10:role_type, 11:chief_type
    const rows = staffData.map(s => {
        return [
            s.id, // A
            s.name, // B
            '', // C
            s.order, // D
            '', // E
            '', // F
            '', // G
            '', // H
            '', // I
            '', // J
            '', // K
            ''  // L
        ];
    });

    if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 12).setValues(rows);
    }

    console.log('Staff List Updated. Count: ' + rows.length);
}
