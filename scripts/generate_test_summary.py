#!/usr/bin/env python3
"""
测试汇总报告生成脚本
整合单元测试、集成测试、性能测试等所有测试结果，生成统一的测试报告
"""

import os
import sys
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
import re
from typing import Dict, List, Any


class TestSummaryGenerator:
    """测试汇总报告生成器"""
    
    def __init__(self, reports_dir: str = "reports"):
        self.reports_dir = Path(reports_dir)
        self.reports_dir.mkdir(exist_ok=True)
        self.summary_data = {
            "generated_at": datetime.now().isoformat(),
            "overall_status": "UNKNOWN",
            "test_results": {},
            "coverage": {},
            "performance": {},
            "security": {},
            "quality_metrics": {}
        }
    
    def collect_unit_test_results(self) -> Dict[str, Any]:
        """收集单元测试结果"""
        unit_report = self.reports_dir / "unit_test_report.html"
        results = {
            "status": "NOT_RUN",
            "total": 0,
            "passed": 0,
            "failed": 0,
            "errors": 0,
            "skipped": 0,
            "pass_rate": 0.0,
            "execution_time": "N/A"
        }
        
        if unit_report.exists():
            try:
                with open(unit_report, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 解析HTML报告中的测试统计
                passed_match = re.search(r'(\d+) passed', content)
                failed_match = re.search(r'(\d+) failed', content)
                error_match = re.search(r'(\d+) error', content)
                skipped_match = re.search(r'(\d+) skipped', content)
                time_match = re.search(r'(\d+\.?\d*) seconds', content)
                
                if passed_match:
                    results["passed"] = int(passed_match.group(1))
                if failed_match:
                    results["failed"] = int(failed_match.group(1))
                if error_match:
                    results["errors"] = int(error_match.group(1))
                if skipped_match:
                    results["skipped"] = int(skipped_match.group(1))
                if time_match:
                    results["execution_time"] = f"{time_match.group(1)}s"
                
                results["total"] = results["passed"] + results["failed"] + results["errors"] + results["skipped"]
                
                if results["total"] > 0:
                    results["pass_rate"] = (results["passed"] / results["total"]) * 100
                    results["status"] = "PASSED" if results["failed"] == 0 and results["errors"] == 0 else "FAILED"
                else:
                    results["status"] = "NO_TESTS"
                    
            except Exception as e:
                print(f"⚠️ 解析单元测试报告失败: {e}")
                results["status"] = "ERROR"
        
        return results
    
    def collect_integration_test_results(self) -> Dict[str, Any]:
        """收集集成测试结果"""
        integration_report = self.reports_dir / "integration_test_report.html"
        results = {
            "status": "NOT_RUN",
            "total": 0,
            "passed": 0,
            "failed": 0,
            "pass_rate": 0.0,
            "execution_time": "N/A"
        }
        
        if integration_report.exists():
            try:
                with open(integration_report, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                passed_match = re.search(r'(\d+) passed', content)
                failed_match = re.search(r'(\d+) failed', content)
                time_match = re.search(r'(\d+\.?\d*) seconds', content)
                
                if passed_match:
                    results["passed"] = int(passed_match.group(1))
                if failed_match:
                    results["failed"] = int(failed_match.group(1))
                if time_match:
                    results["execution_time"] = f"{time_match.group(1)}s"
                
                results["total"] = results["passed"] + results["failed"]
                
                if results["total"] > 0:
                    results["pass_rate"] = (results["passed"] / results["total"]) * 100
                    results["status"] = "PASSED" if results["failed"] == 0 else "FAILED"
                else:
                    results["status"] = "NO_TESTS"
                    
            except Exception as e:
                print(f"⚠️ 解析集成测试报告失败: {e}")
                results["status"] = "ERROR"
                
        return results
    
    def collect_performance_test_results(self) -> Dict[str, Any]:
        """收集性能测试结果"""
        performance_report = self.reports_dir / "performance_test_report.html"
        results = {
            "status": "NOT_RUN",
            "benchmarks": [],
            "summary": "N/A"
        }
        
        if performance_report.exists():
            try:
                with open(performance_report, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 简单解析性能测试结果
                if "PASSED" in content:
                    results["status"] = "PASSED"
                elif "FAILED" in content:
                    results["status"] = "FAILED"
                else:
                    results["status"] = "UNKNOWN"
                    
                results["summary"] = "性能测试已执行，详细结果请查看完整报告"
                
            except Exception as e:
                print(f"⚠️ 解析性能测试报告失败: {e}")
                results["status"] = "ERROR"
                
        return results
    
    def collect_coverage_data(self) -> Dict[str, Any]:
        """收集代码覆盖率数据"""
        coverage_xml = self.reports_dir / "coverage.xml"
        results = {
            "line_coverage": 0.0,
            "branch_coverage": 0.0,
            "total_lines": 0,
            "covered_lines": 0,
            "status": "NOT_AVAILABLE"
        }
        
        if coverage_xml.exists():
            try:
                tree = ET.parse(coverage_xml)
                root = tree.getroot()
                
                # 解析覆盖率信息
                coverage_elem = root.find(".//coverage")
                if coverage_elem is not None:
                    line_rate = float(coverage_elem.get("line-rate", 0))
                    branch_rate = float(coverage_elem.get("branch-rate", 0))
                    
                    results["line_coverage"] = line_rate * 100
                    results["branch_coverage"] = branch_rate * 100
                    results["status"] = "AVAILABLE"
                    
                    # 计算总行数和覆盖行数
                    for package in root.findall(".//package"):
                        lines_valid = int(package.get("lines-valid", 0))
                        lines_covered = int(package.get("lines-covered", 0))
                        results["total_lines"] += lines_valid
                        results["covered_lines"] += lines_covered
                        
            except Exception as e:
                print(f"⚠️ 解析覆盖率报告失败: {e}")
                results["status"] = "ERROR"
                
        return results
    
    def collect_security_data(self) -> Dict[str, Any]:
        """收集安全扫描数据"""
        security_json = self.reports_dir / "security_report.json"
        results = {
            "status": "NOT_RUN",
            "high_issues": 0,
            "medium_issues": 0,
            "low_issues": 0,
            "total_issues": 0,
            "summary": "未执行安全扫描"
        }
        
        if security_json.exists():
            try:
                with open(security_json, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                if 'results' in data:
                    for result in data['results']:
                        severity = result.get('issue_severity', 'LOW')
                        if severity == 'HIGH':
                            results["high_issues"] += 1
                        elif severity == 'MEDIUM':
                            results["medium_issues"] += 1
                        else:
                            results["low_issues"] += 1
                    
                    results["total_issues"] = results["high_issues"] + results["medium_issues"] + results["low_issues"]
                    
                    if results["high_issues"] == 0:
                        results["status"] = "PASSED"
                        results["summary"] = "未发现高危安全问题"
                    else:
                        results["status"] = "FAILED"
                        results["summary"] = f"发现{results['high_issues']}个高危安全问题"
                        
            except Exception as e:
                print(f"⚠️ 解析安全报告失败: {e}")
                results["status"] = "ERROR"
                
        return results
    
    def generate_html_report(self) -> str:
        """生成HTML格式的汇总报告"""
        unit_results = self.summary_data["test_results"].get("unit_tests", {})
        integration_results = self.summary_data["test_results"].get("integration_tests", {})
        performance_results = self.summary_data["test_results"].get("performance_tests", {})
        coverage_data = self.summary_data["coverage"]
        security_data = self.summary_data["security"]
        
        html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI校对助手 - 测试汇总报告</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
        .header {{ text-align: center; margin-bottom: 30px; border-bottom: 2px solid #007acc; padding-bottom: 20px; }}
        .status-badge {{ display: inline-block; padding: 5px 15px; border-radius: 20px; color: white; font-weight: bold; margin: 5px; }}
        .status-passed {{ background-color: #28a745; }}
        .status-failed {{ background-color: #dc3545; }}
        .status-warning {{ background-color: #ffc107; color: #212529; }}
        .status-info {{ background-color: #17a2b8; }}
        .metrics-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 30px; }}
        .metric-card {{ background: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 4px solid #007acc; }}
        .metric-title {{ font-size: 1.2em; font-weight: bold; margin-bottom: 10px; color: #333; }}
        .metric-value {{ font-size: 2em; font-weight: bold; color: #007acc; }}
        .metric-detail {{ margin-top: 10px; color: #666; }}
        .progress-bar {{ width: 100%; height: 20px; background-color: #e9ecef; border-radius: 10px; overflow: hidden; margin: 10px 0; }}
        .progress-fill {{ height: 100%; background-color: #28a745; transition: width 0.3s ease; }}
        .test-summary {{ margin-bottom: 30px; }}
        .test-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        .test-table th, .test-table td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
        .test-table th {{ background-color: #f8f9fa; font-weight: bold; }}
        .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #666; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🤖 AI校对助手 - 测试汇总报告</h1>
            <p>生成时间: {self.summary_data['generated_at']}</p>
            <div class="status-badge status-{self.summary_data['overall_status'].lower()}">{self.summary_data['overall_status']}</div>
        </div>
        
        <div class="metrics-grid">
            <div class="metric-card">
                <div class="metric-title">📊 单元测试</div>
                <div class="metric-value">{unit_results.get('pass_rate', 0):.1f}%</div>
                <div class="metric-detail">
                    通过: {unit_results.get('passed', 0)} / 总计: {unit_results.get('total', 0)}<br>
                    状态: <span class="status-badge status-{unit_results.get('status', 'unknown').lower()}">{unit_results.get('status', 'UNKNOWN')}</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: {unit_results.get('pass_rate', 0)}%"></div>
                </div>
            </div>
            
            <div class="metric-card">
                <div class="metric-title">🔗 集成测试</div>
                <div class="metric-value">{integration_results.get('pass_rate', 0):.1f}%</div>
                <div class="metric-detail">
                    通过: {integration_results.get('passed', 0)} / 总计: {integration_results.get('total', 0)}<br>
                    状态: <span class="status-badge status-{integration_results.get('status', 'unknown').lower()}">{integration_results.get('status', 'UNKNOWN')}</span>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: {integration_results.get('pass_rate', 0)}%"></div>
                </div>
            </div>
            
            <div class="metric-card">
                <div class="metric-title">📈 代码覆盖率</div>
                <div class="metric-value">{coverage_data.get('line_coverage', 0):.1f}%</div>
                <div class="metric-detail">
                    覆盖行数: {coverage_data.get('covered_lines', 0)} / {coverage_data.get('total_lines', 0)}<br>
                    分支覆盖: {coverage_data.get('branch_coverage', 0):.1f}%
                </div>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: {coverage_data.get('line_coverage', 0)}%"></div>
                </div>
            </div>
            
            <div class="metric-card">
                <div class="metric-title">🔒 安全扫描</div>
                <div class="metric-value">{security_data.get('high_issues', 0)}</div>
                <div class="metric-detail">
                    高危问题数<br>
                    中危: {security_data.get('medium_issues', 0)} | 低危: {security_data.get('low_issues', 0)}<br>
                    状态: <span class="status-badge status-{security_data.get('status', 'unknown').lower()}">{security_data.get('status', 'UNKNOWN')}</span>
                </div>
            </div>
        </div>
        
        <div class="test-summary">
            <h2>📋 详细测试结果</h2>
            <table class="test-table">
                <thead>
                    <tr>
                        <th>测试类型</th>
                        <th>状态</th>
                        <th>通过/总计</th>
                        <th>通过率</th>
                        <th>执行时间</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>单元测试</td>
                        <td><span class="status-badge status-{unit_results.get('status', 'unknown').lower()}">{unit_results.get('status', 'UNKNOWN')}</span></td>
                        <td>{unit_results.get('passed', 0)}/{unit_results.get('total', 0)}</td>
                        <td>{unit_results.get('pass_rate', 0):.1f}%</td>
                        <td>{unit_results.get('execution_time', 'N/A')}</td>
                    </tr>
                    <tr>
                        <td>集成测试</td>
                        <td><span class="status-badge status-{integration_results.get('status', 'unknown').lower()}">{integration_results.get('status', 'UNKNOWN')}</span></td>
                        <td>{integration_results.get('passed', 0)}/{integration_results.get('total', 0)}</td>
                        <td>{integration_results.get('pass_rate', 0):.1f}%</td>
                        <td>{integration_results.get('execution_time', 'N/A')}</td>
                    </tr>
                    <tr>
                        <td>性能测试</td>
                        <td><span class="status-badge status-{performance_results.get('status', 'unknown').lower()}">{performance_results.get('status', 'UNKNOWN')}</span></td>
                        <td>-</td>
                        <td>-</td>
                        <td>-</td>
                    </tr>
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>🚀 AI校对助手项目 | 生成于 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p>详细报告请查看各模块的专项测试报告</p>
        </div>
    </div>
</body>
</html>
        """
        
        return html_content
    
    def generate_markdown_report(self) -> str:
        """生成Markdown格式的汇总报告"""
        unit_results = self.summary_data["test_results"].get("unit_tests", {})
        integration_results = self.summary_data["test_results"].get("integration_tests", {})
        coverage_data = self.summary_data["coverage"]
        security_data = self.summary_data["security"]
        
        status_emoji = {
            "PASSED": "✅",
            "FAILED": "❌", 
            "WARNING": "⚠️",
            "NOT_RUN": "⏸️",
            "ERROR": "💥"
        }
        
        markdown_content = f"""# 🤖 AI校对助手 - 测试汇总报告

**生成时间**: {self.summary_data['generated_at']}  
**整体状态**: {status_emoji.get(self.summary_data['overall_status'], '❓')} {self.summary_data['overall_status']}

---

## 📊 测试概览

| 测试类型 | 状态 | 通过/总计 | 通过率 | 执行时间 |
|---------|------|----------|--------|----------|
| 单元测试 | {status_emoji.get(unit_results.get('status'), '❓')} {unit_results.get('status', 'UNKNOWN')} | {unit_results.get('passed', 0)}/{unit_results.get('total', 0)} | {unit_results.get('pass_rate', 0):.1f}% | {unit_results.get('execution_time', 'N/A')} |
| 集成测试 | {status_emoji.get(integration_results.get('status'), '❓')} {integration_results.get('status', 'UNKNOWN')} | {integration_results.get('passed', 0)}/{integration_results.get('total', 0)} | {integration_results.get('pass_rate', 0):.1f}% | {integration_results.get('execution_time', 'N/A')} |

## 📈 质量指标

### 代码覆盖率
- **行覆盖率**: {coverage_data.get('line_coverage', 0):.1f}%
- **分支覆盖率**: {coverage_data.get('branch_coverage', 0):.1f}%
- **覆盖行数**: {coverage_data.get('covered_lines', 0)} / {coverage_data.get('total_lines', 0)}

### 安全扫描
- **状态**: {status_emoji.get(security_data.get('status'), '❓')} {security_data.get('status', 'UNKNOWN')}
- **高危问题**: {security_data.get('high_issues', 0)}
- **中危问题**: {security_data.get('medium_issues', 0)}
- **低危问题**: {security_data.get('low_issues', 0)}

## 📋 详细结果

### 单元测试结果
- ✅ 通过: {unit_results.get('passed', 0)}
- ❌ 失败: {unit_results.get('failed', 0)}
- 💥 错误: {unit_results.get('errors', 0)}
- ⏭️ 跳过: {unit_results.get('skipped', 0)}

### 集成测试结果
- ✅ 通过: {integration_results.get('passed', 0)}
- ❌ 失败: {integration_results.get('failed', 0)}

---

## 📁 相关报告文件

- `unit_test_report.html` - 单元测试详细报告
- `integration_test_report.html` - 集成测试详细报告
- `coverage_html/index.html` - 代码覆盖率详细报告
- `security_report.json` - 安全扫描详细报告

---

*报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*
"""
        
        return markdown_content
    
    def determine_overall_status(self):
        """确定整体状态"""
        unit_status = self.summary_data["test_results"].get("unit_tests", {}).get("status", "NOT_RUN")
        integration_status = self.summary_data["test_results"].get("integration_tests", {}).get("status", "NOT_RUN")
        security_status = self.summary_data["security"].get("status", "NOT_RUN")
        coverage_rate = self.summary_data["coverage"].get("line_coverage", 0)
        
        # 判断整体状态
        if unit_status == "FAILED" or integration_status == "FAILED" or security_status == "FAILED":
            self.summary_data["overall_status"] = "FAILED"
        elif coverage_rate < 80:  # 覆盖率低于80%视为警告
            self.summary_data["overall_status"] = "WARNING"
        elif unit_status == "PASSED" and integration_status == "PASSED" and security_status == "PASSED":
            self.summary_data["overall_status"] = "PASSED"
        elif unit_status == "NOT_RUN" and integration_status == "NOT_RUN":
            self.summary_data["overall_status"] = "NOT_RUN"
        else:
            self.summary_data["overall_status"] = "WARNING"
    
    def generate_summary(self):
        """生成完整的测试汇总报告"""
        print("🚀 开始生成测试汇总报告...")
        
        # 收集各类测试结果
        print("📊 收集单元测试结果...")
        self.summary_data["test_results"]["unit_tests"] = self.collect_unit_test_results()
        
        print("🔗 收集集成测试结果...")
        self.summary_data["test_results"]["integration_tests"] = self.collect_integration_test_results()
        
        print("📈 收集性能测试结果...")
        self.summary_data["test_results"]["performance_tests"] = self.collect_performance_test_results()
        
        print("📊 收集代码覆盖率数据...")
        self.summary_data["coverage"] = self.collect_coverage_data()
        
        print("🔒 收集安全扫描数据...")
        self.summary_data["security"] = self.collect_security_data()
        
        # 确定整体状态
        self.determine_overall_status()
        
        # 生成报告文件
        print("📝 生成HTML汇总报告...")
        html_report = self.generate_html_report()
        html_file = self.reports_dir / "test_summary_report.html"
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_report)
        
        print("📝 生成Markdown汇总报告...")
        markdown_report = self.generate_markdown_report()
        markdown_file = self.reports_dir / "test_summary_report.md"
        with open(markdown_file, 'w', encoding='utf-8') as f:
            f.write(markdown_report)
        
        # 保存JSON数据
        json_file = self.reports_dir / "test_summary_data.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.summary_data, f, ensure_ascii=False, indent=2)
        
        # 输出汇总信息
        print("\n" + "="*50)
        print("📋 测试汇总报告生成完成!")
        print(f"📊 整体状态: {self.summary_data['overall_status']}")
        print(f"📁 HTML报告: {html_file}")
        print(f"📁 Markdown报告: {markdown_file}")
        print(f"📁 JSON数据: {json_file}")
        print("="*50)
        
        return self.summary_data


def main():
    """主函数"""
    generator = TestSummaryGenerator()
    
    try:
        summary_data = generator.generate_summary()
        
        # 根据整体状态决定退出码
        if summary_data["overall_status"] == "FAILED":
            print("❌ 测试整体状态：失败")
            sys.exit(1)
        elif summary_data["overall_status"] == "WARNING":
            print("⚠️ 测试整体状态：警告")
            sys.exit(0)  # 警告不影响构建
        else:
            print("✅ 测试整体状态：成功")
            sys.exit(0)
            
    except Exception as e:
        print(f"💥 生成测试汇总报告失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 