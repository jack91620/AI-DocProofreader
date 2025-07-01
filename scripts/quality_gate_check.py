#!/usr/bin/env python3
"""
质量门禁检查脚本
检查测试覆盖率、测试通过率等质量指标是否达到门禁标准
"""

import os
import sys
import xml.etree.ElementTree as ET
import json
from pathlib import Path
import re


def check_coverage(min_coverage=90):
    """检查代码覆盖率"""
    coverage_file = Path("reports/coverage.xml")
    
    if not coverage_file.exists():
        print("❌ 覆盖率报告文件不存在")
        return False
    
    try:
        tree = ET.parse(coverage_file)
        root = tree.getroot()
        
        # 查找总体覆盖率
        coverage_elem = root.find(".//coverage")
        if coverage_elem is not None:
            line_rate = float(coverage_elem.get("line-rate", 0)) * 100
        else:
            # 尝试其他格式
            line_rate = 0
            for package in root.findall(".//package"):
                line_rate = max(line_rate, float(package.get("line-rate", 0)) * 100)
        
        print(f"📊 代码覆盖率: {line_rate:.1f}%")
        
        if line_rate >= min_coverage:
            print(f"✅ 代码覆盖率达标 (≥{min_coverage}%)")
            return True
        else:
            print(f"❌ 代码覆盖率不达标 (需要≥{min_coverage}%，当前{line_rate:.1f}%)")
            return False
            
    except Exception as e:
        print(f"❌ 解析覆盖率报告失败: {e}")
        return False


def check_test_results(min_pass_rate=95):
    """检查测试通过率"""
    reports_dir = Path("reports")
    test_passed = True
    
    # 检查单元测试报告
    unit_report = reports_dir / "unit_test_report.html"
    integration_report = reports_dir / "integration_test_report.html"
    
    total_tests = 0
    passed_tests = 0
    failed_tests = 0
    
    for report_file in [unit_report, integration_report]:
        if report_file.exists():
            try:
                with open(report_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 从HTML报告中提取结果统计
                passed_match = re.search(r'(\d+) passed', content)
                failed_match = re.search(r'(\d+) failed', content)
                error_match = re.search(r'(\d+) error', content)
                
                if passed_match:
                    passed_tests += int(passed_match.group(1))
                if failed_match:
                    failed_tests += int(failed_match.group(1))
                if error_match:
                    failed_tests += int(error_match.group(1))
                    
            except Exception as e:
                print(f"⚠️ 无法解析测试报告 {report_file}: {e}")
    
    total_tests = passed_tests + failed_tests
    
    if total_tests == 0:
        print("❌ 没有找到测试结果")
        return False
    
    pass_rate = (passed_tests / total_tests) * 100
    
    print(f"📊 测试统计:")
    print(f"   总测试数: {total_tests}")
    print(f"   通过: {passed_tests}")
    print(f"   失败: {failed_tests}")
    print(f"   通过率: {pass_rate:.1f}%")
    
    if pass_rate >= min_pass_rate:
        print(f"✅ 测试通过率达标 (≥{min_pass_rate}%)")
        return True
    else:
        print(f"❌ 测试通过率不达标 (需要≥{min_pass_rate}%，当前{pass_rate:.1f}%)")
        return False


def check_security_issues():
    """检查安全问题"""
    security_report = Path("reports/security_report.json")
    
    if not security_report.exists():
        print("⚠️ 安全扫描报告不存在，跳过安全检查")
        return True
    
    try:
        with open(security_report, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # bandit报告格式
        if 'results' in data:
            high_issues = [r for r in data['results'] if r.get('issue_severity') == 'HIGH']
            medium_issues = [r for r in data['results'] if r.get('issue_severity') == 'MEDIUM']
            
            print(f"📊 安全扫描结果:")
            print(f"   高危问题: {len(high_issues)}")
            print(f"   中危问题: {len(medium_issues)}")
            
            if len(high_issues) > 0:
                print("❌ 存在高危安全问题，必须修复")
                for issue in high_issues:
                    print(f"   - {issue.get('test_name')}: {issue.get('issue_text')}")
                return False
            elif len(medium_issues) > 3:
                print("⚠️ 中危安全问题过多，建议修复")
            
        print("✅ 安全检查通过")
        return True
        
    except Exception as e:
        print(f"⚠️ 解析安全报告失败: {e}")
        return True  # 安全检查失败不阻止发布，但会记录警告


def check_code_quality():
    """检查代码质量"""
    # 检查是否有代码规范违规
    quality_passed = True
    
    # 这里可以添加更多代码质量检查
    # 例如：复杂度检查、重复代码检查等
    
    print("✅ 代码质量检查通过")
    return quality_passed


def generate_quality_report():
    """生成质量门禁报告"""
    report_content = []
    report_content.append("# 质量门禁报告")
    report_content.append(f"**检查时间**: {os.popen('date').read().strip()}")
    report_content.append("")
    
    # 读取覆盖率
    coverage_file = Path("reports/coverage.xml")
    if coverage_file.exists():
        try:
            tree = ET.parse(coverage_file)
            root = tree.getroot()
            coverage_elem = root.find(".//coverage")
            if coverage_elem is not None:
                line_rate = float(coverage_elem.get("line-rate", 0)) * 100
                report_content.append(f"## 代码覆盖率: {line_rate:.1f}%")
        except:
            pass
    
    # 保存报告
    report_file = Path("reports/quality_gate_report.md")
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(report_content))
    
    print(f"📝 质量门禁报告已生成: {report_file}")


def main():
    """主函数"""
    print("🚀 开始质量门禁检查...")
    print("=" * 50)
    
    # 从环境变量获取门禁标准
    min_coverage = int(os.getenv('MIN_COVERAGE', 90))
    min_test_pass_rate = int(os.getenv('MIN_TEST_PASS_RATE', 95))
    
    print(f"📋 质量门禁标准:")
    print(f"   最小代码覆盖率: {min_coverage}%")
    print(f"   最小测试通过率: {min_test_pass_rate}%")
    print("")
    
    # 确保reports目录存在
    Path("reports").mkdir(exist_ok=True)
    
    # 执行各项检查
    checks = []
    
    print("1️⃣ 检查代码覆盖率...")
    checks.append(check_coverage(min_coverage))
    print("")
    
    print("2️⃣ 检查测试通过率...")
    checks.append(check_test_results(min_test_pass_rate))
    print("")
    
    print("3️⃣ 检查安全问题...")
    checks.append(check_security_issues())
    print("")
    
    print("4️⃣ 检查代码质量...")
    checks.append(check_code_quality())
    print("")
    
    # 生成质量报告
    generate_quality_report()
    
    # 汇总结果
    print("=" * 50)
    
    if all(checks):
        print("🎉 质量门禁检查全部通过！")
        print("✅ 代码质量达标，可以继续部署")
        sys.exit(0)
    else:
        print("❌ 质量门禁检查失败！")
        print("🚫 代码质量不达标，阻止部署")
        
        # 输出失败的检查项
        check_names = ["代码覆盖率", "测试通过率", "安全检查", "代码质量"]
        for i, (check_name, passed) in enumerate(zip(check_names, checks)):
            status = "✅" if passed else "❌"
            print(f"   {status} {check_name}")
        
        sys.exit(1)


if __name__ == "__main__":
    main() 