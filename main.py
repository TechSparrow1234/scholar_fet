from scholarly import scholarly
import pandas as pd
import re
import time
import os

def search_scholars_by_research_areas(n, research_areas):
    all_scholars_info = []
    total_areas = len(research_areas)
    
    for index, area in enumerate(research_areas, start=1):
        print(f"开始搜索研究方向 {area}（第 {index} 个方向，共 {total_areas} 个方向）...")
        
        try:
            search_query = scholarly.search_keyword(area)
            found_authors = 0
            
            for i, author in enumerate(search_query):
                if i >= n:
                    break
                try:
                    print(f"正在获取研究方向 {area} 下第 {i + 1} 个学者的信息...")
                    filled_author = scholarly.fill(author, sections=["basics", "indices", "publications"])
                    
                    scholar_info = {
                        "name": filled_author.get('name', 'N/A'),
                        "affiliation": filled_author.get('affiliation', 'N/A'),
                        "research_areas": ', '.join(filled_author.get('interests', [])),
                        "papers": []
                    }
                    
                    for j, pub in enumerate(filled_author.get('publications', [])):
                        if j >= 2:
                            break
                        try:
                            print(f"正在获取学者 {scholar_info['name']} 的第 {j + 1} 篇论文信息...")
                            filled_pub = scholarly.fill(pub) 

                            # 直接提取 `pub_year`
                            paper_info = {
                                "title": filled_pub.get('bib', {}).get('title', 'N/A'),
                                "year": filled_pub.get('bib', {}).get('pub_year', '未知'),
                                "citations": filled_pub.get('num_citations', 'N/A')
                            }
                            scholar_info["papers"].append(paper_info)
                            time.sleep(1)  # 控制爬取速度
                        except Exception as e:
                            print(f"获取论文信息时出错: {e}")

                    all_scholars_info.append(scholar_info)
                    found_authors += 1
                except Exception as e:
                    print(f"获取学者信息时出错: {e}")
            
            print(f"研究方向 {area} 搜索完成，共找到 {found_authors} 个学者。")
        except Exception as e:
            print(f"搜索 {area} 相关学者时出错: {e}")
    
    return all_scholars_info



# 在终端输出收集到的学者信息
def print_scholars_info(scholars_info):
    for scholar in scholars_info:
        print(f"姓名: {scholar['name']}")
        print(f"学院: {scholar['affiliation']}")
        print(f"研究方向: {scholar['research_areas']}")
        print("论文信息:")
        for paper in scholar['papers']:
            print(f"  - 标题: {paper['title']}")
            print(f"    年份: {paper['year']}")
            print(f"    引用数: {paper['citations']}")
        print("-" * 50)



def save_to_excel(scholars_info, file_name="scholars_info.xlsx"):
    print("开始将学者信息保存到 Excel 文件...")

    # 确保有学者信息
    if not scholars_info:
        print("未找到任何学者信息，Excel 文件不会被创建。")
        return

    # 检查 Pandas 依赖
    try:
        import openpyxl
    except ImportError:
        print("缺少 openpyxl 库，请运行 'pip install openpyxl' 进行安装。")
        return

    # 处理数据
    data = []
    added_data = set()

    for scholar in scholars_info:
        # 存储论文的标题、年份和引用数
        paper_titles = []
        paper_years = []
        paper_citations = []

        for paper in scholar["papers"]:
            title = re.sub(r'[\n\r\t]', '', paper["title"]) if isinstance(paper["title"], str) else paper["title"]
            year = re.sub(r'[\n\r\t]', '', str(paper["year"])) if isinstance(paper["year"], (str, int, float)) else paper["year"]
            citations = re.sub(r'[\n\r\t]', '', str(paper["citations"])) if isinstance(paper["citations"], (str, int, float)) else paper["citations"]

            paper_titles.append(title)
            paper_years.append(year)
            paper_citations.append(citations)

        # 每个学者占一行，将论文信息合并
        row = {
            "姓名": scholar["name"],
            "单位": scholar["affiliation"],
            "研究方向": scholar["research_areas"],
            "论文1标题": paper_titles[0] if len(paper_titles) > 0 else 'N/A',
            "论文1年份": paper_years[0] if len(paper_years) > 0 else 'N/A',
            "论文1引用数": paper_citations[0] if len(paper_citations) > 0 else 'N/A',
            "论文2标题": paper_titles[1] if len(paper_titles) > 1 else 'N/A',
            "论文2年份": paper_years[1] if len(paper_years) > 1 else 'N/A',
            "论文2引用数": paper_citations[1] if len(paper_citations) > 1 else 'N/A'
        }
        data.append(row)

    df = pd.DataFrame(data)

    # 直接保存到当前目录
    try:
        df.to_excel(file_name, index=False)
        print(f"学者信息已成功保存到当前目录: {file_name}。")
    except Exception as e:
        print(f"保存 Excel 文件时出错: {e}")




# 示例研究方向列表
n = int(input("请输入你想搜集的学者数量："))
research_areas = ["Agent-Based Modeling"]
print("开始搜索所有研究方向的学者信息...")
scholars_info = search_scholars_by_research_areas(n, research_areas)
print("所有研究方向的学者信息搜索完成。")

# 输出获取到的学者信息
print("获取到的学者信息如下：")
print_scholars_info(scholars_info)

# 保存信息到 Excel 文件
save_to_excel(scholars_info, "scholars_info.xlsx")