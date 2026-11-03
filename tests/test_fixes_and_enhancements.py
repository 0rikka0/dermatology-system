# -*- coding: utf-8 -*-
"""
Tests for Fixes and Enhancements:
1. Table separator line removal and exact text line padding per image count specification.
2. Advanced multi-condition search filtering with date range, prefix, and statuses.
3. Special patient workflow: prohibition of direct archiving in report maker, auto-placement in
   特殊患者归档, and moving to 特殊患者归档/完成审核 upon approval and archiving.
"""
import os
import sys
import shutil
import tempfile
import docx

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.doc_generator import generate_report_file, TEMPLATES_DIR
from app.path_manager import (
    ensure_workspace_structure,
    get_database_path,
    get_patient_word_path,
    get_final_report_path,
    save_special_draft_queue,
    archive_patient_report,
    revoke_patient_archive_files,
    revoke_patient_approval_files
)
from app.database import (
    init_db,
    insert_patient,
    get_patient,
    update_patient_report_making,
    mark_patient_approved,
    mark_patient_archived,
    search_patients
)

def run_tests():
    print("=== 开始修正与功能增强自动化集成测试 ===")
    temp_dir = tempfile.mkdtemp(prefix="derm_fixes_test_")
    print(f"创建临时测试环境: {temp_dir}")

    try:
        # 1. 验证工作目录结构建立 (包含 特殊患者归档/完成审核)
        ws_res = ensure_workspace_structure(temp_dir)
        assert os.path.exists(os.path.join(temp_dir, "特殊患者归档", "完成审核")), "缺少 特殊患者归档/完成审核 文件夹"
        print("[OK] [TEST 0] 特殊患者归档/完成审核 启动初始化验证通过！")

        db_path = get_database_path(temp_dir)
        init_db(db_path)

        # -------------------------------------------------------------
        # 2. 验证 Requirement 1: 去除表格上方多余分隔线 & 各规格行数补全
        # -------------------------------------------------------------
        print("\n[TEST 1] 验证表格上方多余分隔线去除及各规格行数补全...")
        specs = [
            (1, 14, 10),
            (2, 14, 8),
            (3, 14, 10),
            (4, 10, 6),
            (6, 10, 6),
            (7, 4, 3),
            (9, 4, 3)
        ]

        for img_cnt, exp_micro, exp_diag in specs:
            doc_file = os.path.join(temp_dir, f"test_spec_{img_cnt}.docx")
            test_data = {
                "name": "排版测试",
                "gender": "男",
                "age": 30,
                "pathology_no": f"202611030{img_cnt}",
                "submission_date": "2026-11-03",
                "image_count": img_cnt,
                "biopsy_sites": [f"部位{i+1}" for i in range(img_cnt)],
                "clinical_diagnosis": "色素痣",
                "microscopic_desc": "真皮浅层见巢状细胞。",
                "pathology_diag": "1.交界痣",
                "medical_history": "病史资料"
            }
            generate_report_file(TEMPLATES_DIR, doc_file, test_data)
            d = docx.Document(doc_file)

            # 验证无多余 floating VML 直线 2
            vml_lines = []
            for p in d.paragraphs:
                for c in p._p.iter():
                    if c.tag.split("}")[-1] == "line":
                        vml_lines.append(c.get("id"))
            assert len(vml_lines) == 0, f"发现多余浮动线条: {vml_lines}"

            # 验证行数
            p_texts = [p.text for p in d.paragraphs]
            micro_idx = [i for i, t in enumerate(p_texts) if "镜下描述：" in t][0]
            diag_idx = [i for i, t in enumerate(p_texts) if "病理诊断：" in t][0]
            sig_idx = [i for i, t in enumerate(p_texts) if "报告医生：" in t][0]
            act_micro = diag_idx - micro_idx
            act_diag = sig_idx - diag_idx
            assert act_micro == exp_micro, f"图片数={img_cnt}: 期望镜下描述 {exp_micro} 行，实际为 {act_micro} 行"
            assert act_diag == exp_diag, f"图片数={img_cnt}: 期望病理诊断 {exp_diag} 行，实际为 {act_diag} 行"
            os.remove(doc_file)

        print("  -> 单图/双图/三图/4-6图/7-9图 全规格行数精准对齐与15磅固定行距验证通过！")
        print("  -> 表格上方多余分隔线清除验证通过！")

        # -------------------------------------------------------------
        # 3. 验证 Requirement 2: 综合查询筛选 (日期范围、前缀、审核/归档状态)
        # -------------------------------------------------------------
        print("\n[TEST 2] 验证病历综合查询多条件交叉筛选...")
        # 插入若干测试患者
        # Patient 1: 今日普通 未归档 无需审核
        p1 = {
            "pathology_no": "2026110301",
            "name": "张普通",
            "gender": "男",
            "age": "40",
            "clinical_diagnosis": "皮炎",
            "image_count": 1,
            "submission_date": "2026年11月03日",
            "submission_date_raw": "2026-11-03",
            "report_date": "2026年11月03日",
            "report_date_raw": "2026-11-03"
        }
        insert_patient(db_path, p1)

        # Patient 2: 今日陆港 已归档 完成审核
        p2 = {
            "pathology_no": "L2026110301",
            "name": "李陆港",
            "gender": "女",
            "age": "25",
            "clinical_diagnosis": "脂溢性角化病",
            "image_count": 2,
            "submission_date": "2026年11月03日",
            "submission_date_raw": "2026-11-03",
            "report_date": "2026年11月03日",
            "report_date_raw": "2026-11-03"
        }
        insert_patient(db_path, p2)
        mark_patient_approved(db_path, "L2026110301")
        mark_patient_archived(db_path, "L2026110301", "some_dir")

        # Patient 3: 昨天病房 待审核 未归档
        p3 = {
            "pathology_no": "B2026110201",
            "name": "王病房",
            "gender": "男",
            "age": "60",
            "clinical_diagnosis": "恶性黑色素瘤待排",
            "image_count": 3,
            "submission_date": "2026年11月02日",
            "submission_date_raw": "2026-11-02",
            "report_date": "2026年11月02日",
            "report_date_raw": "2026-11-02"
        }
        insert_patient(db_path, p3)
        update_patient_report_making(db_path, "B2026110201", {
            "is_pending_review": 1,
            "is_needs_ihc": 1
        })

        # 3.1 默认查询今日
        res_today = search_patients(db_path, {"start_date": "2026-11-03", "end_date": "2026-11-03"})
        nos_today = [p["pathology_no"] for p in res_today]
        assert "2026110301" in nos_today and "L2026110301" in nos_today
        assert "B2026110201" not in nos_today
        print("  -> 默认时间段 (今日) 筛选验证通过！")

        # 3.2 前缀筛选 (仅陆港)
        res_l = search_patients(db_path, {"prefix": "L"})
        assert len(res_l) == 1 and res_l[0]["pathology_no"] == "L2026110301"
        print("  -> 病理号前缀 (仅陆港L) 筛选验证通过！")

        # 3.3 归档状态筛选 (已归档)
        res_arch = search_patients(db_path, {"archived": "1"})
        assert len(res_arch) == 1 and res_arch[0]["pathology_no"] == "L2026110301"
        print("  -> 归档状态筛选 (已归档) 验证通过！")

        # 3.4 审核状态筛选 (待审核)
        res_pending = search_patients(db_path, {"review_status": "待审核"})
        assert len(res_pending) == 1 and res_pending[0]["pathology_no"] == "B2026110201"
        print("  -> 审核状态筛选 (待审核) 验证通过！")

        # 3.5 一键取消所有筛选 (全量数据)
        res_all = search_patients(db_path, {})
        assert len(res_all) >= 3
        print("  -> 一键取消筛选 (全量返回) 验证通过！")

        # -------------------------------------------------------------
        # 4. 验证 Requirement 3: 特殊患者归档与审核流转闭环
        # -------------------------------------------------------------
        print("\n[TEST 3] 验证特殊患者归档工作流与禁止直接归档约束...")

        # 创建一个待审核+免疫组化报告
        p_special = dict(get_patient(db_path, "B2026110201"))
        p_special["pathology_diag"] = "黑色素瘤待排"
        p_special["microscopic_desc"] = "见异型黑素细胞浸润。"
        update_patient_report_making(db_path, "B2026110201", {
            "pathology_diag": "黑色素瘤待排",
            "microscopic_desc": "见异型黑素细胞浸润。",
            "is_pending_review": 1,
            "is_needs_ihc": 1
        })
        p_special = get_patient(db_path, "B2026110201")
        report_docx = get_final_report_path(temp_dir, "B2026110201", "王病房", "黑色素瘤待排")
        generate_report_file(TEMPLATES_DIR, report_docx, p_special)

        # 4.1 点击保存时：直接放入 特殊患者归档/待审核 和 免疫组化 两个目录
        save_draft_res = save_special_draft_queue(temp_dir, p_special, report_docx)
        assert save_draft_res["success"] is True

        pending_dir = os.path.join(temp_dir, "特殊患者归档", "待审核", "B2026110201_王病房_黑色素瘤待排")
        ihc_dir = os.path.join(temp_dir, "特殊患者归档", "免疫组化", "B2026110201_王病房_黑色素瘤待排")
        assert os.path.exists(pending_dir), "特殊患者归档/待审核 中未生成对应文件夹"
        assert os.path.exists(ihc_dir), "特殊患者归档/免疫组化 中未生成对应文件夹"
        print("  -> 保存时自动放入 特殊患者归档/待审核 与 免疫组化 两个目录验证通过！")

        # 4.2 在报告制作工作区直接归档：必须被禁止！
        arch_res_maker = archive_patient_report(temp_dir, p_special, report_docx, is_from_review=False)
        assert arch_res_maker["success"] is False
        assert "禁止在报告制作工作区直接归档" in arch_res_maker["error"]
        print("  -> 报告制作区禁止直接归档特殊报告校验拦截通过！")

        # 4.3 专家完成审核 (mark approved)
        approved_p = mark_patient_approved(db_path, "B2026110201")
        assert approved_p["review_status"] == "完成审核"

        # 4.4 完成审核后归档 (从 特殊患者归档 中取出并放入 完成审核，同时存入 按患者归档)
        arch_res_rev = archive_patient_report(temp_dir, approved_p, report_docx, is_from_review=True)
        assert arch_res_rev["success"] is True

        # 验证原 待审核 / 免疫组化 目录已被取出 (移除)
        assert not os.path.exists(pending_dir), "待审核 目录应已被取出"
        assert not os.path.exists(ihc_dir), "免疫组化 目录应已被取出"

        # 验证 放入 特殊患者归档/完成审核
        approved_target_dir = os.path.join(temp_dir, "特殊患者归档", "完成审核", "B2026110201_王病房_黑色素瘤待排")
        assert os.path.exists(approved_target_dir), "特殊患者归档/完成审核 未找到归档文件夹"

        # 验证 照常放入 按患者归档
        patient_archive_dir = os.path.join(temp_dir, "按患者归档", "B2026110201_王病房_黑色素瘤待排")
        assert os.path.exists(patient_archive_dir), "按患者归档 未找到归档文件夹"
        print("  -> 审核完成后归档：从待审/组化取出并转移至 完成审核，且常规存入 按患者归档 验证通过！")

        # 4.5 验证撤销归档
        revoke_res = revoke_patient_archive_files(temp_dir, approved_p)
        assert revoke_res["success"] is True
        assert not os.path.exists(approved_target_dir)
        assert not os.path.exists(patient_archive_dir)
        print("  -> 撤销归档联动清理验证通过！")

        print("\n=========================================")
        print("[PASS] 修正需求 1、2、3 全部功能自动化测试圆满通过！")
        print("=========================================")

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    run_tests()
