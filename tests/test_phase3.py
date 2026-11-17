# -*- coding: utf-8 -*-
"""
End-to-End Test Suite for Phase 3:
- 3.1 Report Output Workspace (Regular vs. Special Approved lists, prefix filters, batch print, revoke archive & approval)
- 3.2 Signing Sheet Generation (Date range, multiple prefix selection, candidate querying, Excel formatting, preview & printing)
"""
import os
import sys
import shutil
import tempfile
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from app.path_manager import (
    ensure_workspace_structure,
    get_database_path,
    get_patient_photo_dir,
    get_final_report_path,
    archive_patient_report,
    find_patient_archive_dir,
    revoke_patient_archive_files,
    revoke_patient_approval_files,
    get_signing_sheet_dir
)
from app.database import (
    init_db,
    insert_patient,
    get_patient,
    mark_patient_archived,
    mark_patient_approved,
    get_output_regular_patients,
    get_output_special_approved_patients,
    revoke_patient_archive_db,
    revoke_patient_approval_db,
    get_patients_for_signing_sheet
)
from app.doc_generator import generate_report_file
from app.signing_sheet_generator import generate_signing_sheet_excel, render_signing_sheet_html

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "模板")

def run_tests():
    print("=== 开始 Phase 3: 报告输出与签名单自动化集成测试 ===")
    test_dir = tempfile.mkdtemp(prefix="derm_phase3_test_")
    print(f"创建临时测试工作目录: {test_dir}")

    try:
        ensure_workspace_structure(test_dir)
        db_path = get_database_path(test_dir)
        init_db(db_path)
        print("✓ 工作目录与数据库初始化通过")

        # 1. 准备4位测试患者 (覆盖普通号、L分院号、B病房号、T特殊号，以及常规vs待审组化)
        p1_data = {
            "pathology_no": "2026110301",
            "prefix": "",
            "name": "普通张",
            "gender": "男",
            "age": "30",
            "clinical_diagnosis": "皮内痣",
            "submission_date": "2026年11月03日",
            "submission_date_raw": "2026-11-03",
            "report_date_raw": "2026-11-03",
            "report_doctor": "张医生",
            "review_doctor": "李主任",
            "is_pending_review": 0,
            "is_needs_ihc": 0,
            "review_status": "无",
            "archived": 1,
            "created_at": "2026-11-03 09:00:00"
        }
        p2_data = {
            "pathology_no": "L2026110301",
            "prefix": "L",
            "name": "分院李",
            "gender": "女",
            "age": "42",
            "clinical_diagnosis": "脂溢性角化",
            "submission_date": "2026年11月03日",
            "submission_date_raw": "2026-11-03",
            "report_date_raw": "2026-11-03",
            "report_doctor": "张医生",
            "review_doctor": "李主任",
            "is_pending_review": 0,
            "is_needs_ihc": 0,
            "review_status": "无",
            "archived": 1,
            "created_at": "2026-11-03 09:30:00"
        }
        p3_data = {
            "pathology_no": "B2026110301",
            "prefix": "B",
            "name": "病房王",
            "gender": "男",
            "age": "58",
            "clinical_diagnosis": "基底细胞癌",
            "submission_date": "2026年11月03日",
            "submission_date_raw": "2026-11-03",
            "report_date_raw": "2026-11-03",
            "report_doctor": "张医生",
            "review_doctor": "李主任",
            "is_pending_review": 1,
            "is_needs_ihc": 1,
            "review_status": "完成审核",
            "archived": 1,
            "created_at": "2026-11-03 10:00:00"
        }
        p4_data = {
            "pathology_no": "T2026110201",
            "prefix": "T",
            "name": "特殊赵",
            "gender": "女",
            "age": "65",
            "clinical_diagnosis": "血管肉瘤待排",
            "submission_date": "2026年11月02日",
            "submission_date_raw": "2026-11-02",
            "report_date_raw": "2026-11-02",
            "report_doctor": "张医生",
            "review_doctor": "李主任",
            "is_pending_review": 1,
            "is_needs_ihc": 0,
            "review_status": "完成审核",
            "archived": 1,
            "created_at": "2026-11-02 15:00:00"
        }

        # 录入数据库
        insert_patient(db_path, p1_data)
        insert_patient(db_path, p2_data)
        insert_patient(db_path, p3_data)
        insert_patient(db_path, p4_data)

        # 生成假报告与归档文件夹
        for p in [p1_data, p2_data, p3_data, p4_data]:
            diag = p["clinical_diagnosis"]
            final_p = get_final_report_path(test_dir, p["pathology_no"], p["name"], diag)
            generate_report_file(TEMPLATES_DIR, final_p, p)
            archive_res = archive_patient_report(test_dir, p, final_p)
            mark_patient_archived(db_path, p["pathology_no"], archive_res["archive_dir"])
            if p["review_status"] == "完成审核":
                mark_patient_approved(db_path, p["pathology_no"])

        # 模拟 p4 为昨日报告
        import sqlite3
        with sqlite3.connect(db_path) as conn:
            conn.execute("UPDATE patients SET created_at = '2026-11-02 15:00:00', report_date_raw = '2026-11-02', submission_date_raw = '2026-11-02' WHERE pathology_no = 'T2026110201'")
            conn.commit()

        print("✓ 测试患者数据录入、最终报告生成与打包归档就绪")

        # ---------------- TEST 3.1: 报告输出工作区双列表与前缀过滤 ----------------
        print("\n[TEST 3.1] 验证报告输出工作区列表与前缀过滤...")
        # 列表1: 常规已归档报告 (应有 p1 普通张 和 p2 分院李)
        reg_all = get_output_regular_patients(db_path, prefix="ALL")
        assert len(reg_all) == 2, f"常规报告列表总数应为2，实际为: {len(reg_all)}"
        reg_nos = [p["pathology_no"] for p in reg_all]
        assert "2026110301" in reg_nos and "L2026110301" in reg_nos
        print("  -> 常规已归档报告总数与成员验证通过")

        # 前缀过滤测试
        reg_none = get_output_regular_patients(db_path, prefix="NONE")
        assert len(reg_none) == 1 and reg_none[0]["pathology_no"] == "2026110301", "仅普通筛选异常"
        reg_l = get_output_regular_patients(db_path, prefix="L")
        assert len(reg_l) == 1 and reg_l[0]["pathology_no"] == "L2026110301", "仅分院(L)筛选异常"
        reg_b = get_output_regular_patients(db_path, prefix="B")
        assert len(reg_b) == 0, "B病房常规报告应为空"
        print("  -> 常规报告前缀过滤 (ALL, NONE, L, B) 验证通过！")

        # 列表2: 特殊审核完成报告 (应有 p3 病房王 和 p4 特殊赵)
        spec_all = get_output_special_approved_patients(db_path, prefix="ALL")
        assert len(spec_all) == 2, f"特殊已审核列表总数应为2，实际为: {len(spec_all)}"
        spec_nos = [p["pathology_no"] for p in spec_all]
        assert "B2026110301" in spec_nos and "T2026110201" in spec_nos
        print("  -> 特殊审核完成报告总数与成员验证通过")

        spec_b = get_output_special_approved_patients(db_path, prefix="B")
        assert len(spec_b) == 1 and spec_b[0]["pathology_no"] == "B2026110301", "仅病房(B)筛选异常"
        spec_t = get_output_special_approved_patients(db_path, prefix="T")
        assert len(spec_t) == 1 and spec_t[0]["pathology_no"] == "T2026110201", "仅特殊(T)筛选异常"
        print("  -> 特殊报告前缀过滤 (B, T) 验证通过！")

        # 今日报告过滤测试
        spec_today = get_output_special_approved_patients(db_path, prefix="ALL", today_only=True, today_str="2026-11-03")
        assert len(spec_today) == 1 and spec_today[0]["pathology_no"] == "B2026110301", "今日特殊报告过滤异常"
        print("  -> 今日报告专属过滤验证通过！")

        # ---------------- TEST 3.1.1: 归档撤销与审核撤销 ----------------
        print("\n[TEST 3.1.1] 验证撤销归档与撤销审核联动操作...")
        # 1. 撤销 p1 的归档
        revoke_patient_archive_files(test_dir, p1_data)
        p1_revoked = revoke_patient_archive_db(db_path, "2026110301")
        assert p1_revoked["archived"] == 0
        p1_arch_dir = find_patient_archive_dir(test_dir, "2026110301", "普通张")
        assert not p1_arch_dir or not os.path.exists(p1_arch_dir), "按患者归档目录应被移除"
        # 撤销后应自动从输出列表中移出
        reg_after_revoke = get_output_regular_patients(db_path, prefix="ALL")
        assert len(reg_after_revoke) == 1 and reg_after_revoke[0]["pathology_no"] == "L2026110301"
        print("  -> 撤销归档及文件同步清理验证通过！")

        # 2. 撤销 p3 的完成审核状态
        p3_revoked_rev = revoke_patient_approval_db(db_path, "B2026110301")
        assert p3_revoked_rev["review_status"] == "待审核/免疫组化", f"审核状态应退回，实际为: {p3_revoked_rev['review_status']}"
        revoke_patient_approval_files(test_dir, p3_data, p3_revoked_rev["review_status"])
        spec_after_revoke = get_output_special_approved_patients(db_path, prefix="ALL")
        assert len(spec_after_revoke) == 1 and spec_after_revoke[0]["pathology_no"] == "T2026110201", "撤销审核后应移出特殊审核完成列表"
        print("  -> 撤销审核状态及特殊归档恢复验证通过！")

        # ---------------- TEST 3.2: 签名单候选患者拉取与多条件筛选 ----------------
        print("\n[TEST 3.2] 验证报告领取签名单范围筛选与候选拉取...")
        # 全量拉取 (不看归档或审核状态)
        all_signing = get_patients_for_signing_sheet(db_path, prefixes=["ALL"])
        assert len(all_signing) == 4, f"全量候选应为4人，实际为: {len(all_signing)}"
        print("  -> 无视归档/审核状态全量拉取验证通过！")

        # 时间范围筛选
        date_filtered = get_patients_for_signing_sheet(db_path, start_date="2026-11-03", end_date="2026-11-03", prefixes=["ALL"])
        assert len(date_filtered) == 3, f"11月3日报告应有3人，实际为: {len(date_filtered)}"
        print("  -> 时间范围筛选 (起始~结束) 验证通过！")

        # 多前缀复选框组合筛选 (如同时选 [普通, 分院(L)])
        multi_pfx = get_patients_for_signing_sheet(db_path, prefixes=["NONE", "L"])
        assert len(multi_pfx) == 2
        multi_pfx_nos = [p["pathology_no"] for p in multi_pfx]
        assert "2026110301" in multi_pfx_nos and "L2026110301" in multi_pfx_nos
        print("  -> 多前缀复选组合筛选 (普通 + 分院) 验证通过！")

        # ---------------- TEST 3.2.1: Excel 签名单生成与格式校验 ----------------
        print("\n[TEST 3.2.1] 验证 Excel 签名单生成与格式规范...")
        gen_result = generate_signing_sheet_excel(
            test_dir,
            all_signing,
            start_date="2026-11-02",
            end_date="2026-11-03"
        )
        assert gen_result["success"] is True
        excel_path = gen_result["file_path"]
        assert os.path.exists(excel_path), f"签名单文件缺失: {excel_path}"
        assert os.path.dirname(excel_path) == os.path.join(test_dir, "领取表"), "签名单必须保存在 工作目录/领取表 中"

        # 检验生成的 Excel 格式严格符合要求
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        # 表头
        headers = [ws.cell(1, c).value for c in range(1, 4)]
        assert headers == ["病理号", "姓名", "签名"], f"表头不符: {headers}"
        
        # 数据填充校验
        row2_vals = [ws.cell(2, 1).value, ws.cell(2, 2).value]
        assert row2_vals == ["2026110301", "普通张"], f"第2行数据不符: {row2_vals}"
        
        # 单元格边框校验 (全表格细线边框)
        c2 = ws.cell(2, 1)
        assert c2.border.top.style == "thin" and c2.border.bottom.style == "thin", "单元格边框样式缺失"
        assert c2.alignment.horizontal == "center", "居中对齐设置缺失"
        print(f"  -> 生成签名单Excel验证通过: {os.path.basename(excel_path)} (包含 {len(all_signing)} 人)")

        # HTML 预览生成校验
        html_prev = render_signing_sheet_html(all_signing, date_range_str="2026-11-02 至 2026-11-03")
        assert "某三甲医院皮肤科" in html_prev
        assert "普通张" in html_prev and "2026110301" in html_prev
        print("  -> HTML签名单高保真预览渲染验证通过！")

        print("=========================================")
        print("🎉 Phase 3 (3.1 报告输出与 3.2 签名单生成) 所有功能全部通过！")
        print("=========================================")

    finally:
        shutil.rmtree(test_dir, ignore_errors=True)

if __name__ == "__main__":
    run_tests()
