from scholarly import scholarly
import pandas as pd
import re
import time


def search_scholars_by_research_areas(research_areas):
    all_scholars_info = []
    total_areas = len(research_areas)
    for index, area in enumerate(research_areas, start=1):
        print(f"开始搜索研究方向 {area}（第 {index} 个，共 {total_areas} 个）...")
        try:
            # 搜索与该研究方向相关的学者
            search_query = scholarly.search_keyword(area)
            found_authors = 0
            for i, author in enumerate(search_query):
                if i >= 100:  # 每个研究方向只取前 3 个学者
                    break
                try:
                    print(f"正在获取研究方向 {area} 下第 {i + 1} 个学者的信息...")
                    # 获取学者的详细信息，包括基本信息、索引信息和共同作者信息
                    filled_author = scholarly.fill(author, sections=["basics", "indices", "publications"])
                    scholar_info = {
                        "name": filled_author.get('name', 'N/A'),
                        "affiliation": filled_author.get('affiliation', 'N/A'),
                        "position": "N/A",  # scholarly 库未直接提供职称信息
                        "research_areas": ', '.join(filled_author.get('interests', [])),
                        "papers": []
                    }
                    # 获取学者的前 2 篇论文信息
                    for j, pub in enumerate(filled_author.get('publications', [])):
                        if j >= 2:
                            break
                        try:
                            print(f"正在获取学者 {scholar_info['name']} 的第 {j + 1} 篇论文信息...")
                            filled_pub = scholarly.fill(pub)  # 这里可能卡住
                            paper_info = {
                                "title": filled_pub.get('bib', {}).get('title', 'N/A'),
                                "year": filled_pub.get('bib', {}).get('year', 'N/A'),
                                "citations": filled_pub.get('num_citations', 'N/A')
                            }
                            scholar_info["papers"].append(paper_info)
                        except Exception as e:
                            print(f"获取论文信息时出错: {e}")
                        time.sleep(1)
                    all_scholars_info.append(scholar_info)
                    found_authors += 1
                except Exception as e:
                    print(f"获取学者信息时出错: {e}")
            print(f"研究方向 {area} 搜索完成，共找到 {found_authors} 个学者。")
        except Exception as e:
            print(f"搜索 {area} 相关学者时出错: {e}")
    return all_scholars_info


def print_scholars_info(scholars_info):
    for scholar in scholars_info:
        print(f"姓名: {scholar['name']}")
        print(f"学院: {scholar['affiliation']}")
        print(f"职称: {scholar['position']}")
        print(f"研究方向: {scholar['research_areas']}")
        print("论文信息:")
        for paper in scholar['papers']:
            print(f"  - 标题: {paper['title']}")
            print(f"    年份: {paper['year']}")
            print(f"    引用数: {paper['citations']}")
        print("-" * 50)



def save_to_excel(scholars_info, file_name):
    print("开始将学者信息保存到 Excel 文件...")
    data = []
    added_data = set()

    for scholar in scholars_info:
        for paper in scholar["papers"]:
            # 数据清洗，去除特殊字符
            title = re.sub(r'[\n\r\t]', '', paper["title"]) if isinstance(paper["title"], str) else paper["title"]
            year = re.sub(r'[\n\r\t]', '', str(paper["year"])) if isinstance(paper["year"], (str, int, float)) else paper["year"]
            citations = re.sub(r'[\n\r\t]', '', str(paper["citations"])) if isinstance(paper["citations"], (str, int, float)) else paper["citations"]

            key = (scholar["name"], title)
            if key not in added_data:
                row = {
                    "姓名": scholar["name"],
                    "学院": scholar["affiliation"],
                    "职称": scholar["position"],
                    "研究方向": scholar["research_areas"],
                    "论文标题": title,
                    "论文年份": year,
                    "论文引用数": citations
                }
                data.append(row)
                added_data.add(key)

    df = pd.DataFrame(data)

    # 文件格式兼容性处理
    if file_name.endswith('.xlsx'):
        df.to_excel(file_name, index=False)
    elif file_name.endswith('.xls'):
        df.to_excel(file_name, index=False, engine='xlwt')
    else:
        raise ValueError("Unsupported file format. Only.xlsx and.xls are supported.")

    print(f"学者信息已成功保存到 {file_name} 文件。")


# 示例研究方向列表
research_areas = ["Agent-Based Modeling"]
print("开始搜索所有研究方向的学者信息...")
scholars_info = search_scholars_by_research_areas(research_areas)
print("所有研究方向的学者信息搜索完成。")

# 输出获取到的学者信息
print("获取到的学者信息如下：")
print_scholars_info(scholars_info)

# 保存信息到 Excel 文件
save_to_excel(scholars_info, "scholars_info.xlsx")