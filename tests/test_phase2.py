# -*- coding: utf-8 -*-
"""
End-to-End Test Suite for Phase 2: Report Making, Review, Template Engine, and Archiving.
"""
import os
import sys
import shutil
import tempfile
import json
import docx

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from app.path_manager import (
    ensure_workspace_structure,
    get_database_path,
    get_patient_photo_dir,
    find_clinical_photo_folder,
    get_final_report_path,
    archive_patient_report,
    save_special_draft_queue,
    sync_review_overwrites,
    locate_patient_all_paths,
    delete_patient_all_files,
    sync_status_file_operations
)
from app.database import (
    init_db,
    insert_patient,
    get_patient,
    update_patient_report_making,
    mark_patient_archived,
    mark_patient_approved,
    get_patients_for_report_making,
    get_patients_for_review,
    update_patient_status,
    delete_patient
)
from app.template_manager import (
    load_text_templates,
    search_templates,
    recommend_templates,
    parse_photo_filename,
    format_pathology_diagnosis
)
from app.doc_generator import generate_report_file
from app.doc_preview import render_preview_html

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "模板")

def run_tests():
    print("=== 开始 Phase 2 自动化集成测试 ===")
    test_dir = tempfile.mkdtemp(prefix="derm_phase2_test_")
    print(f"创建临时测试工作目录: {test_dir}")

    try:
        # 1. 结构与数据库初始化
        ensure_workspace_structure(test_dir)
        db_path = get_database_path(test_dir)
        init_db(db_path)
        print("✓ 工作目录结构与SQLite数据库初始化成功")

        # 2. 模板解析与诊断合成测试
        all_tmpls = load_text_templates()
        assert len(all_tmpls) >= 200, f"模板数量不足: {len(all_tmpls)}"
        print(f"✓ 成功加载文字模板: 共 {len(all_tmpls)} 篇")

        # 诊断提取测试
        p1 = parse_photo_filename("1.2.交界痣.jpg")
        assert p1["indices"] == ["1", "2"] and p1["diagnosis"] == "交界痣"
        p2 = parse_photo_filename("3.皮内痣.png")
        assert p2["indices"] == ["3"] and p2["diagnosis"] == "皮内痣"
        
        diag_formatted = format_pathology_diagnosis([
            (p1["indices"], p1["diagnosis"]),
            (p2["indices"], p2["diagnosis"])
        ])
        assert diag_formatted == "1.2.交界痣；3.皮内痣", f"诊断合成不符合预期: {diag_formatted}"
        print(f"✓ 显微镜照片命名解析与病理诊断合成验证通过: {diag_formatted}")

        # 模板推荐测试
        recs = recommend_templates(["交界痣", "皮内痣"])
        assert len(recs) > 0, "推荐模板未找到"
        assert any("交界痣" in r["name"] for r in recs), "未能推荐交界痣模板"
        print(f"✓ 智能推荐文字模板成功匹配: {[r['name'] for r in recs[:3]]}")

        # 3. 创建测试患者
        patient_data = {
            "pathology_no": "2026110301",
            "name": "张三",
            "gender": "男",
            "age": "35",
            "inpatient_no": "ZY1001",
            "outpatient_no": "MZ2002",
            "referring_doctor": "赵医生",
            "clinical_diagnosis": "色素痣",
            "image_count": 3,
            "biopsy_sites": ["左胸前", "左胸中", "左胸后"],
            "medical_history": "胸部肿物3个月余，逐渐增大。",
            "submission_date": "2026年11月03日",
            "submission_date_raw": "2026-11-03",
            "report_doctor": "张医生",
            "review_doctor": "李主任",
            "word_path": os.path.join(test_dir, "统一归档", "病理报告录入", "2026110301_张三.docx"),
            "photo_dir": get_patient_photo_dir(test_dir, "2026110301", "张三")
        }
        os.makedirs(os.path.dirname(patient_data["word_path"]), exist_ok=True)
        generate_report_file(TEMPLATES_DIR, patient_data["word_path"], patient_data)
        insert_patient(db_path, patient_data)
        print("✓ 患者信息录入及草稿Word报告生成成功")

        # 4. 模拟输入显微镜图片和临床照片
        photo_dir = patient_data["photo_dir"]
        os.makedirs(photo_dir, exist_ok=True)
        # 用占位图复制模拟真实显微镜图片
        sample_img = os.path.join(TEMPLATES_DIR, "placeholder.jpeg")
        img1_path = os.path.join(photo_dir, "1.2.交界痣.jpg")
        img2_path = os.path.join(photo_dir, "3.皮内痣.jpg")
        shutil.copy2(sample_img, img1_path)
        shutil.copy2(sample_img, img2_path)

        clin_dir = os.path.join(test_dir, "统一归档", "临床照片", "张三")
        os.makedirs(clin_dir, exist_ok=True)
        clin_img = os.path.join(clin_dir, "clinical_front.jpg")
        shutil.copy2(sample_img, clin_img)
        print("✓ 模拟显微镜照片与临床照片就绪")

        # 5. 临床照片查找逻辑验证
        lookup = find_clinical_photo_folder(test_dir, "张三", "2026110301")
        assert lookup["found"] is True and lookup["path"] == clin_dir, "未能找到患者临床照片目录"
        print(f"✓ 临床照片检索成功: {lookup['path']}")

        # 6. 报告制作与保存 (2.7)
        final_doc_path = get_final_report_path(test_dir, "2026110301", "张三", diag_formatted)
        report_making_data = dict(patient_data)
        report_making_data.update({
            "pathology_diag": diag_formatted,
            "microscopic_desc": "1.2.表皮大致正常，真表皮交界处可见痣细胞巢。\n3.真皮内可见大量痣细胞巢，成熟现象明显。",
            "selected_images": [
                {"slot": 1, "path": img1_path, "site": "左胸前"},
                {"slot": 2, "path": img1_path, "site": "左胸中"},
                {"slot": 3, "path": img2_path, "site": "左胸后"}
            ],
            "image_count": 3,
            "final_report_path": final_doc_path,
            "is_pending_review": 1,
            "is_needs_ihc": 1
        })

        generate_report_file(TEMPLATES_DIR, final_doc_path, report_making_data)
        assert os.path.exists(final_doc_path), "最终输出报告文件未生成"
        
        # 验证生成的最终Word排版
        doc = docx.Document(final_doc_path)
        doc_texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        assert any("镜下描述：" in t for t in doc_texts)
        assert any("1.2.表皮大致正常" in t for t in doc_texts)
        # Requirement 1: “病理诊断：”小标题后换行写诊断内容
        assert any(t == "病理诊断：" for t in doc_texts), "病理诊断：小标题应单独成段"
        assert any("1.2.交界痣；3.皮内痣" in t for t in doc_texts), "病理诊断内容应紧随其后"
        assert any("报告医生：张医生" in t for t in doc_texts)
        assert any("病史：" in t for t in doc_texts)

        # 更新数据库并同步保存至特殊待审队列 (Requirement 3)
        update_patient_report_making(db_path, "2026110301", report_making_data)
        p_db = get_patient(db_path, "2026110301")
        assert p_db["pathology_diag"] == diag_formatted
        assert p_db["is_pending_review"] == 1
        assert p_db["is_needs_ihc"] == 1
        assert p_db["review_status"] == "待审核/免疫组化"
        print(f"✓ 报告制作与保存成功: {os.path.basename(final_doc_path)}")

        # 保存时自动放入特殊患者归档 (待审核与免疫组化两个目录)
        save_special_draft_queue(test_dir, p_db, final_doc_path)
        safe_name = f"2026110301_张三_{diag_formatted}"
        pending_arch = os.path.join(test_dir, "特殊患者归档", "待审核", safe_name)
        ihc_arch = os.path.join(test_dir, "特殊患者归档", "免疫组化", safe_name)
        assert os.path.exists(pending_arch), "待审核特殊归档缺失"
        assert os.path.exists(ihc_arch), "免疫组化特殊归档缺失"

        # 制作区直接归档应被禁止 (Requirement 3)
        block_res = archive_patient_report(test_dir, p_db, final_doc_path, is_from_review=False)
        assert block_res["success"] is False

        # 8. 报告审核工作区测试 (2.9)
        review_list = get_patients_for_review(db_path, include_approved=False)
        assert len(review_list) == 1 and review_list[0]["pathology_no"] == "2026110301"

        # 模拟审核修改医生
        report_making_data["report_doctor"] = "王医生"
        report_making_data["review_doctor"] = "李主任"
        generate_report_file(TEMPLATES_DIR, final_doc_path, report_making_data)

        # 审核完成标记与归档 (Requirement 3: 从特殊待审取出，放入完成审核与按患者归档)
        p_approved = mark_patient_approved(db_path, "2026110301")
        archive_res = archive_patient_report(test_dir, p_approved, final_doc_path, is_from_review=True)
        assert archive_res["success"] is True
        main_arch = archive_res["archive_dir"]
        assert os.path.exists(main_arch)
        assert os.path.exists(os.path.join(main_arch, "临床照片", "clinical_front.jpg"))
        assert os.path.exists(os.path.join(main_arch, "病理照片", "1.2.交界痣.jpg"))
        assert any(f.endswith(".docx") for f in os.listdir(main_arch))

        # 验证从待审核/免疫组化中取出，放入完成审核
        assert not os.path.exists(pending_arch)
        assert not os.path.exists(ihc_arch)
        assert os.path.exists(os.path.join(test_dir, "特殊患者归档", "完成审核", safe_name))
        
        mark_patient_archived(db_path, "2026110301", main_arch)
        p_arch = get_patient(db_path, "2026110301")
        assert p_arch["archived"] == 1
        print(f"✓ 档案打包归档及特殊患者同步归档成功: {main_arch}")

        # 同步覆盖验证
        sync_res = sync_review_overwrites(test_dir, p_arch, final_doc_path)
        assert sync_res["success"] is True
        assert len(sync_res["overwritten"]) >= 2
        print(f"✓ 审核修改保存并自动覆盖各归档目录: 共覆盖 {len(sync_res['overwritten'])} 处")

        # 验证覆盖后文件中医生名称
        doc_overwritten = docx.Document(final_doc_path)
        sig_p = [p.text for p in doc_overwritten.paragraphs if "报告医生：" in p.text][0]
        assert "王医生" in sig_p and "李主任" in sig_p, f"医生签名未更新: {sig_p}"
        print(f"✓ 覆盖后签名验证通过: {sig_p.strip()}")

        # 完成审核
        mark_patient_approved(db_path, "2026110301")
        p_approved = get_patient(db_path, "2026110301")
        assert p_approved["review_status"] == "完成审核"
        assert p_approved["is_pending_review"] == 0

        # 默认不显示已审核患者测试
        review_list_hidden = get_patients_for_review(db_path, include_approved=False)
        assert len(review_list_hidden) == 0, "完成审核患者未隐藏"
        review_list_shown = get_patients_for_review(db_path, include_approved=True)
        assert len(review_list_shown) == 1, "开启显示后未能看到已完成审核患者"
        print("✓ 完成审核标记及状态切换验证通过！")

        # 9. 重新保存已归档报告测试 (Requirement 2: archived 自动变为 2 [需要重新归档])
        re_save_data = dict(report_making_data)
        re_save_data["microscopic_desc"] += "\n镜下补充观察：未见明显恶变迹象。"
        update_patient_report_making(db_path, "2026110301", re_save_data)
        p_re_saved = get_patient(db_path, "2026110301")
        assert p_re_saved["archived"] == 2, f"已归档报告重新保存后 archived 应为 2 (需要重新归档)，实际为: {p_re_saved['archived']}"
        print("✓ 已归档报告再次保存排版时状态正确置为【需要重新归档 (2)】")

        # 10. 关联路径检索测试 (Requirement 3: 综合查询中展开各关联目录)
        all_paths = locate_patient_all_paths(test_dir, p_re_saved)
        path_keys = {p["key"]: p for p in all_paths}
        assert path_keys["word_draft"]["exists"] is True
        assert path_keys["word_final"]["exists"] is True
        assert path_keys["photo_dir"]["exists"] is True
        assert path_keys["archive_dir"]["exists"] is True
        print(f"✓ 关联路径检索正常: 发现 {len(all_paths)} 项关联目录与文件")

        # 11. 状态管理与文件同步测试 (Requirement 3: 修改状态同时撤销/进行文件操作)
        status_sync_res = sync_status_file_operations(test_dir, p_re_saved, new_review_status="无", new_archived_status=0)
        assert status_sync_res["success"] is True
        update_patient_status(db_path, "2026110301", "无", 0)
        p_status_reset = get_patient(db_path, "2026110301")
        assert p_status_reset["review_status"] == "无"
        assert p_status_reset["archived"] == 0
        print("✓ 状态管理修改及关联归档文件撤销同步测试通过")

        # 12. 删除报告及所有关联目录测试 (Requirement 3: 移除数据库条目和所有对应目录)
        del_files_res = delete_patient_all_files(test_dir, p_status_reset)
        assert del_files_res["success"] is True
        delete_patient(db_path, "2026110301")
        p_deleted = get_patient(db_path, "2026110301")
        assert p_deleted is None, "数据库中患者记录未能成功删除"
        assert not os.path.exists(p_status_reset["photo_dir"]), "患者照片目录未能彻底删除"
        print("✓ 删除报告、移除数据库条目及清除所有对应文件目录测试通过！")

        print("=========================================")
        print("🎉 Phase 2 & 最新修改需求所有功能模块与测试用例全部通过！")
        print("=========================================")

    finally:
        shutil.rmtree(test_dir, ignore_errors=True)

if __name__ == "__main__":
    run_tests()
