import time
import json
import uuid
import logging
import sys
import traceback
from chinese_calendar import is_holiday
from datetime import datetime, timedelta
from datetime import timezone
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# 配置日志
logging.basicConfig(filename='error_log.txt', level=logging.ERROR)


def uid():
    """生成唯一标识符"""
    return str(uuid.uuid4())


def format_building_name(building_name):
    """将原始建筑名称格式化为正确名称"""
    # 处理常见前缀，如"济-"
    if building_name.startswith("济-"):
        # 从"济-"后面提取楼号部分
        building_number = building_name[2:]
        return f"山东体育学院{building_number}"
    # 可以添加其他前缀的处理逻辑
    return building_name


def save_courses_to_json(courses, filename='courses.json'):
    """将课程信息保存为JSON格式"""
    formatted_courses = []
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

    for course in courses:
        day_name = weekdays[course['day'] - 1] if 1 <= course['day'] <= 7 else f"星期{course['day']}"
        # 修改这行，正确显示节次范围
        sections_str = f"第{course['sections'][0]}-{course['sections'][1]}节"
        weeks_str = f"第{','.join(str(w) for w in course['weeks'])}周"

        # 修改位置信息的格式化方式，将\n替换为-
        position = course['position'].replace("\n", "-")

        formatted_course = {
            "name": course['name'],
            "teacher": course['teacher'],
            "time": day_name,
            "sections": sections_str,
            "weeks": weeks_str,
            "weeks_array": course['weeks'],
            "position": position,
            "day": course['day'],
            "section_array": course['sections']
        }
        formatted_courses.append(formatted_course)

    # 写入JSON文件
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(formatted_courses, f, ensure_ascii=False, indent=4)

    print(f"课表信息已保存到JSON文件: {filename}")


def get_course_table_html(driver):
    """获取课表HTML内容，处理可能的iframe情况"""
    try:
        # 等待页面完全加载
        time.sleep(3)
        print("页面标题:", driver.title)

        # 检查是否存在iframe
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            print(f"找到 {len(iframes)} 个iframe，尝试切换...")
            for index, iframe in enumerate(iframes):
                try:
                    driver.switch_to.frame(iframe)
                    print(f"已切换到iframe {index + 1}")

                    # 尝试找表格
                    tables = driver.find_elements(By.TAG_NAME, "table")
                    if tables:
                        print(f"在iframe {index + 1}中找到 {len(tables)} 个表格")
                        # 假设第一个表格是课表
                        return driver.page_source

                    driver.switch_to.default_content()
                except Exception as e:
                    print(f"切换到iframe {index + 1}出错: {e}")
                    driver.switch_to.default_content()

        # 如果没有在iframe中找到，则在主页面中寻找
        tables = driver.find_elements(By.TAG_NAME, "table")
        if tables:
            print(f"在主页面找到 {len(tables)} 个表格")
            return driver.page_source

        # 实在找不到表格，尝试查找其他课表相关元素
        divs = driver.find_elements(By.CLASS_NAME, "divOneClass")
        if divs:
            print(f"找到 {len(divs)} 个课程元素")
            return driver.page_source

        # 返回整个页面内容以供分析
        print("未找到具体课表元素，返回整个页面内容")
        return driver.page_source
    except Exception as e:
        print(f"获取课表HTML失败: {e}")
        # 保存页面源码以便调试
        with open("page_source.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("已将页面源码保存至page_source.html文件")
        return driver.page_source  # 返回页面内容，以便后续尝试解析


def parse_weeks_string(weeks_str):
    """将周数字符串转换为周数数组，对应stringToWeeksArray函数"""
    weeks_str = weeks_str.replace('周', '')
    try:
        start, end = map(int, weeks_str.split('-'))
        return list(range(start, end + 1))
    except ValueError:
        # 处理单周情况
        try:
            return [int(weeks_str)]
        except:
            return []


def sections_to_array(section):
    """将节数转换为数组，对应sectionsToArray函数"""
    section = int(section)
    return [section, section + 1]


def parse_course_info(html_content):
    """解析课表HTML并提取课程信息，对应scheduleHtmlParser函数"""
    soup = BeautifulSoup(html_content, 'html.parser')
    course_infos = []

    course_elements = soup.select('.divOneClass')

    for course_element in course_elements:
        parent_td = course_element.find_parent('td')
        row = parent_td.get('row')
        col = parent_td.get('col')

        # 获取并格式化建筑物名称
        building_name = course_element.select_one('.spBuilding').text.strip()
        formatted_building = format_building_name(building_name)
        classroom = course_element.select_one('.spClassroom').text.strip()

        course_info = {
            'name': course_element.select_one('.spLUName').text.strip(),
            'teacher': course_element.select_one('.spTeacherName').text.strip(),
            'weeks': parse_weeks_string(course_element.select_one('.spWeekInfo').text.strip()),
            'position': f"{formatted_building}\n{classroom}",
            'sections': sections_to_array(row),
            'day': int(col)
        }

        # 检查是否有相同 day, weeks 和 sections 的课程
        existing_course = None
        for info in course_infos:
            if (info['day'] == course_info['day'] and
                    info['weeks'] == course_info['weeks'] and
                    info['sections'] == course_info['sections']):
                existing_course = info
                break

        if existing_course:
            # 合并 position
            existing_course['position'] += f"  {course_info['position']}"
        else:
            course_infos.append(course_info)

    return course_infos


def format_and_display_courses(courses):
    """格式化并显示课程信息"""
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

    print("\n===== 课表信息 =====")
    for course in courses:
        day_name = weekdays[course['day'] - 1] if 1 <= course['day'] <= 7 else f"星期{course['day']}"
        # 修改这行，正确显示节次范围
        sections_str = f"第{course['sections'][0]}-{course['sections'][1]}节"
        weeks_str = f"第{','.join(str(w) for w in course['weeks'])}周"

        print(f"\n课程: {course['name']}")
        print(f"教师: {course['teacher']}")
        print(f"时间: {day_name} {sections_str}")
        print(f"周次: {weeks_str}")
        print(f"地点: {course['position']}")

    print("\n===================")


def save_courses_to_file(courses, filename='courses.txt'):
    """保存课程信息到文件"""
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("===== 课表信息 =====\n")

        for course in courses:
            day_name = weekdays[course['day'] - 1] if 1 <= course['day'] <= 7 else f"星期{course['day']}"
            sections_str = f"第{course['sections'][0]}-{course['sections'][1]}节"
            weeks_str = f"第{','.join(str(w) for w in course['weeks'])}周"

            f.write(f"\n课程: {course['name']}\n")
            f.write(f"教师: {course['teacher']}\n")
            f.write(f"时间: {day_name} {sections_str}\n")
            f.write(f"周次: {weeks_str}\n")
            f.write(f"地点: {course['position']}\n")

        f.write("\n===================\n")

    print(f"课表信息已保存到文件: {filename}")


def generate_ics_from_json(json_file, first_week_date=None, alarm_minutes=30):
    """从JSON课表文件生成iCS日历文件，并排除法定节假日"""
    # 加载JSON数据
    with open(json_file, 'r', encoding='utf-8') as f:
        courses = json.load(f)

    if not courses:
        print("课表数据为空，无法生成日历")
        return

    # 确定第一周周一的日期
    if first_week_date is None:
        first_week_date = input("请输入第一周周一的日期 (格式:YYYYMMDD，如20230904): ")

    try:
        initial_time = datetime.strptime(first_week_date, "%Y%m%d")
    except ValueError:
        print("日期格式错误，应为YYYYMMDD")
        return

    # 设置课程时间表 - 每节课的开始和结束时间
    class_timetable = {
        "1": {"startTime": "080000", "endTime": "084000"},
        "2": {"startTime": "085000", "endTime": "093000"},
        "3": {"startTime": "100000", "endTime": "104000"},
        "4": {"startTime": "105000", "endTime": "113000"},
        "5": {"startTime": "133000", "endTime": "141000"},
        "6": {"startTime": "142000", "endTime": "150000"},
        "7": {"startTime": "153000", "endTime": "161000"},
        "8": {"startTime": "162000", "endTime": "170000"},
        "9": {"startTime": "180000", "endTime": "184000"},
        "10": {"startTime": "185000", "endTime": "193000"},
    }

    # 设置提醒时间
    if alarm_minutes > 0:
        a_trigger = f"-PT{alarm_minutes}M"
    else:
        a_trigger = ""

    # 生成日历文件
    utc_now = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    weekdays = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]

    # 写入日历文件头部
    ical_begin_base = f'''BEGIN:VCALENDAR
VERSION:2.0
X-WR-CALNAME:课表
X-APPLE-CALENDAR-COLOR:#FF2968
X-WR-TIMEZONE:Asia/Shanghai
BEGIN:VTIMEZONE
TZID:Asia/Shanghai
X-LIC-LOCATION:Asia/Shanghai
BEGIN:STANDARD
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
TZNAME:CST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE
'''

    output_file = f"课表-{utc_now}.ics"

    try:
        with open(output_file, "w", encoding='UTF-8') as f:
            f.write(ical_begin_base)

            # 处理每门课程
            for course in courses:
                day = course['day']  # 周几（1-7）
                sections = course['section_array']  # 节次数组 [开始节次, 结束节次]
                weeks = course['weeks_array']  # 周数数组 [1,2,3...]

                if not weeks:
                    continue

                # 排序周数，找出开始周和结束周
                weeks.sort()
                start_week = weeks[0]
                end_week = weeks[-1]

                # 检查是单周、双周还是每周
                is_odd = all(w % 2 == 1 for w in weeks)
                is_even = all(w % 2 == 0 for w in weeks)

                if is_odd:
                    week_status = 1  # 单周
                elif is_even:
                    week_status = 2  # 双周
                else:
                    week_status = 0  # 每周

                # 计算课程第一次开始的日期
                delta_time = 7 * (start_week - 1) + day - 1

                if week_status == 1:  # 单周
                    if start_week % 2 == 0:  # 若单周就不变，双周加7
                        delta_time += 7
                elif week_status == 2:  # 双周
                    if start_week % 2 != 0:  # 若双周就不变，单周加7
                        delta_time += 7

                first_time_obj = initial_time + timedelta(days=delta_time)

                # 计算所有可能上课的日期，并检查哪些是节假日
                exclude_dates = []
                for week_num in range(start_week, end_week + 1):
                    # 根据单双周情况判断是否需要排除
                    if week_status == 1 and week_num % 2 == 0:  # 单周，跳过双周
                        continue
                    elif week_status == 2 and week_num % 2 == 1:  # 双周，跳过单周
                        continue

                    # 计算这一周的上课日期
                    week_delta = week_num - start_week
                    class_date = first_time_obj + timedelta(days=7 * week_delta)

                    # 检查是否为法定节假日
                    if is_holiday(class_date):
                        # 生成EXDATE格式的日期字符串
                        exclude_date = class_date.strftime("%Y%m%dT") + class_timetable[str(sections[0])]["startTime"]
                        exclude_dates.append(exclude_date)

                # 构建EXDATE属性
                exdate_str = ""
                if exclude_dates:
                    exdate_str = "EXDATE;TZID=Asia/Shanghai:" + ",".join(exclude_dates) + "\n"

                if week_status == 0:  # 每周
                    extra_status = "1"
                else:
                    extra_status = f'2;BYDAY={weekdays[day - 1]}'

                # 获取课程开始和结束时间 - 修改这部分以正确处理连续两节课
                start_section = str(sections[0])  # 课程开始节次
                end_section = str(sections[1])  # 课程结束节次

                # 使用开始节次的开始时间和结束节次的结束时间
                final_stime_str = first_time_obj.strftime("%Y%m%d") + "T" + class_timetable[start_section]["startTime"]
                final_etime_str = first_time_obj.strftime("%Y%m%d") + "T" + class_timetable[end_section]["endTime"]

                # 计算结束日期
                delta_week = 7 * (end_week - start_week)
                stop_time_obj = first_time_obj + timedelta(days=delta_week + 1)
                stop_time_str = stop_time_obj.strftime("%Y%m%dT%H%M%SZ")

                # 生成提醒部分
                if a_trigger:
                    _alarm_base = f'''BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:This is an event reminder
TRIGGER:{a_trigger}
X-WR-ALARMUID:{uid()}
UID:{uid()}
END:VALARM
'''
                else:
                    _alarm_base = ""

                # 生成事件内容，添加EXDATE排除节假日
                _ical_base = f'''BEGIN:VEVENT
CREATED:{utc_now}
DTSTAMP:{utc_now}
SUMMARY:{course['name']}
DESCRIPTION:教师：{course['teacher']}
LOCATION:{course['position']}
TZID:Asia/Shanghai
SEQUENCE:0
UID:{uid()}
{exdate_str}RRULE:FREQ=WEEKLY;UNTIL={stop_time_str};INTERVAL={extra_status}
DTSTART;TZID=Asia/Shanghai:{final_stime_str}
DTEND;TZID=Asia/Shanghai:{final_etime_str}
X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC
{_alarm_base}END:VEVENT
'''
                f.write(_ical_base)

            # 写入文件尾部
            f.write("\nEND:VCALENDAR")

    except Exception as e:
        print(f"生成iCS文件时出错: {e}")
        return None

    print(f"日历文件已生成: {output_file}")
    return output_file


def main():
    """主函数，整合所有功能"""
    try:
        print("正在初始化浏览器...")

        # 给出所需的url
        url_login_page = "https://jw.sdpei.edu.cn"
        url_kebiao_page = "https://jw.sdpei.edu.cn/Student/CourseTimetable/MyCourseTimeTable.aspx"

        # 启动Edge驱动，开始模拟
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-images")  # 禁用图片加载提高速度
        options.add_experimental_option("detach", True)  # 保持浏览器打开

        # 初始化浏览器
        driver = None
        try:
            driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=options)
        except Exception as e:
            print(f"初始化完成: {e}")
            print("请重新打开该程序")
            input("按回车键退出...")
            return

        print("正在访问教务系统...")
        driver.get(url_login_page)

        # 自动输入账号密码
        username = input("请输入学号: ")
        password = input("请输入密码: ")
        driver.find_element(By.ID, "txtUserName").send_keys(username)
        driver.find_element(By.ID, "txtPassword").send_keys(password)

        # 找到并点击登录按钮，实现登录
        print("正在登录...")
        login_button = driver.find_element(By.ID, "mlbActive")
        actions = ActionChains(driver)
        actions.key_down(Keys.CONTROL).click(login_button).key_up(Keys.CONTROL).perform()

        # 等待新窗口打开
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[-1])

        # 直接访问成绩查询页面
        print("正在访问个人课表页面...")
        driver.get(url_kebiao_page)

        print("正在获取课表信息...")
        html_content = get_course_table_html(driver)

        if html_content:
            print("获取到页面内容，尝试解析...")

            # 保存原始HTML以便调试
            with open("raw_table.html", "w", encoding="utf-8") as f:
                f.write(html_content)
            print("原始HTML已保存到raw_table.html")

            try:
                courses = parse_course_info(html_content)

                if courses:
                    print(f"共解析到 {len(courses)} 门课程")
                    format_and_display_courses(courses)

                    # 询问用户选择保存格式
                    save_option = input("是否保存课表? (y/n): ")
                    if save_option.lower() == 'y':
                        format_option = input("选择保存格式 (1: TXT, 2: JSON, 3: iCS日历文件, 默认JSON): ") or "2"

                        if format_option == "1":
                            filename = input("输入文件名 (默认为 courses.txt): ") or "courses.txt"
                            save_courses_to_file(courses, filename)
                        elif format_option == "3":
                            # 先保存为JSON
                            json_filename = "temp_courses.json"
                            save_courses_to_json(courses, json_filename)
                            # 询问第一周日期
                            first_week = input("请输入第一周周一的日期 (格式:YYYYMMDD，如20230904): ")
                            # 询问是否需要提醒
                            need_alarm = input("是否需要课前提醒? (y/n): ").lower() == 'y'
                            alarm_minutes = 30
                            if need_alarm:
                                try:
                                    alarm_minutes = int(input("请输入提前提醒的分钟数 (默认30): ") or "30")
                                except ValueError:
                                    print("输入错误，使用默认值30分钟")
                                    alarm_minutes = 30
                            else:
                                alarm_minutes = 0
                            # 生成iCS文件
                            ics_file = generate_ics_from_json(json_filename, first_week, alarm_minutes)
                            if ics_file:
                                print(f"iCS日历文件已生成: {ics_file}")
                                print("您可以将此文件导入到手机或电脑的日历应用中")
                        else:
                            filename = input("输入文件名 (默认为 courses.json): ") or "courses.json"
                            save_courses_to_json(courses, filename)
                else:
                    print("未能解析出课程信息，请检查页面内容")

            except Exception as e:
                print(f"解析课表时出错: {e}")
                logging.error(f"解析课表时出错: {e}")
                logging.error(traceback.format_exc())
        else:
            print("获取课表失败，请检查网络连接或登录状态")

        # 关闭浏览器
        if driver:
            driver.quit()

    except Exception as e:
        print(f"程序运行出错: {e}")
        logging.error(f"程序运行出错: {e}")
        logging.error(traceback.format_exc())

    # 防止程序立即关闭
    input("\n程序执行完毕，按回车键退出...")


if __name__ == "__main__":
    main()