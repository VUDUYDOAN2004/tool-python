import sys
from docx import Document

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

def sort_timetable(input_file):
    # Thứ tự chuẩn
    correct_order = [
        ("Thứ hai", "Sáng", ["Toán", "Công nghệ", "Tiếng Việt", "Tiếng Việt"]),
        ("Thứ hai", "Chiều", ["Toán", "Rèn Toán", "HĐTN"]),
        ("Thứ ba", "Sáng", ["Tiếng Việt", "Tiếng Việt", "GDTC", "Toán"]),
        ("Thứ ba", "Chiều", ["Tin học", "TNXH", "Tiếng Anh"]),
        ("Thứ tư", "Sáng", ["Tiếng Anh", "NT(M.T)", "NT(Â.N)", "Đạo đức"]),
        ("Thứ tư", "Chiều", []),  # không có tiết
        ("Thứ năm", "Sáng", ["TNXH", "Tiếng Việt", "Tiếng Việt", "Toán"]),
        ("Thứ năm", "Chiều", ["Tiếng Anh", "Rèn T.Việt", "HĐTN"]),
        ("Thứ sáu", "Sáng", ["Toán", "Tiếng Việt", "GDTC", "Rèn T.Việt"]),
        ("Thứ sáu", "Chiều", ["Tiếng Anh", "Rèn Toán", "HĐTN" ]),
    ]

    doc = Document(input_file)

    # Tìm bảng chính
    table = None
    for t in doc.tables:
        if "Ngày" in t.rows[0].cells[0].text:
            table = t
            break
    if not table:
        print(" Không tìm thấy bảng thời khóa biểu.")
        return

    # Gom tất cả các hàng chứa môn (cột 3 khác rỗng)
    all_rows = [row for row in table.rows[1:] if row.cells[3].text.strip()]

    pos = 0
    for (day, session, subjects) in correct_order:
        size = len(subjects)  # số tiết cần xử lý
        if size == 0:
            print(f"\n {day} - {session}: không có tiết → bỏ qua")
            continue

        rows = all_rows[pos:pos+size]
        if not rows:
            print(f"\n {day} - {session}: không tìm thấy đủ hàng dữ liệu (cần {size})")
            pos += size
            continue

        print(f"\n {day} - {session}")
        print("   Gốc:", [r.cells[3].text.strip() for r in rows])

        # Map dữ liệu gốc
        subj_map = {}
        for r in rows:
            subj = r.cells[3].text.strip()
            subj_map.setdefault(subj, []).append(
                (r.cells[3].text, r.cells[4].text, r.cells[5].text)
            )

        # Tạo dữ liệu mới
        new_data = []
        subj_counter = {}
        for subj in subjects:
            subj_counter[subj] = subj_counter.get(subj, 0)
            if subj in subj_map and subj_counter[subj] < len(subj_map[subj]):
                new_data.append(subj_map[subj][subj_counter[subj]])
                subj_counter[subj] += 1
            else:
                new_data.append((subj, "", ""))

        # Ghi đè lại
        for r, (ns, nl, nm) in zip(rows, new_data):
            r.cells[3].text = ns
            r.cells[4].text = nl
            r.cells[5].text = nm

        print("   Mới:", [ns for ns, _, _ in new_data])
        pos += size

    doc.save(input_file)
    print("\n Hoàn thành, đã chỉnh sửa trực tiếp:", input_file)


if __name__ == "__main__":
    sort_timetable("C:/Users/duydo/Downloads/ĐĂNG KÍ GIẢNG DẠY TUẦN 7.docx")

