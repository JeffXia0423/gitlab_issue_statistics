import gitlab
import xlwt
import datetime

# 用户git账户的token
private_token = 'Sfqsps4z7yhDVXD3JLJo'
# git地址
private_host = 'http://192.168.253.126:8086/'


def getAllProjects():
    client = gitlab.Gitlab(private_host, private_token=private_token)
    projects = client.projects.list(membership=True, all=True)
    return projects


def find_nddl_max(nddl_list):
    nddl_num_max = nddl_list[0]["异常数量"]
    nddl_num_index = 0
    for i in range(1, len(nddl_list)):
        if nddl_num_max < nddl_list[i]["异常数量"]:
            nddl_num_max = nddl_list[i]["异常数量"]
            nddl_num_index = i
    nddl_num_max_list = nddl_list[nddl_num_index]
    del nddl_list[nddl_num_index]
    return nddl_num_max_list


def find_delay_max(assignee_list):
    delay_num_max = assignee_list[0]['逾期数量']
    delay_num_index = 0
    for i in range(1, len(assignee_list)):
        if delay_num_max < assignee_list[i]['逾期数量']:
            delay_num_max = assignee_list[i]['逾期数量']
            delay_num_index = i
    delay_num_max_list = assignee_list[delay_num_index]
    del assignee_list[delay_num_index]
    return delay_num_max_list


def writeExcel(excelPath, data, no_ddl_data):
    workbook = xlwt.Workbook()
    # 获取第一个sheet页
    issue_owner = workbook.add_sheet('责任人统计(逾期+无截止日期)')
    delay_issue = workbook.add_sheet('issue统计(逾期)')
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_LEFT  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER  # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    style = xlwt.XFStyle()  # Create Style
    style.alignment = alignment  # Add Alignment to Style
    delay_issue.col(0).width = 20*256
    delay_issue.col(1).width = 120*256
    delay_issue.col(2).width = 8*256
    delay_issue.col(3).width = 7*256
    delay_issue.col(4).width = 60*256
    row0 = ['项目名称', 'issue标题', '逾期天数', '处理人', '链接']
    assignee_list = []
    for i in range(0, len(row0)):
        delay_issue.write(0, i, row0[i], style)
    for i in range(0, len(data)):
        record = data[i]
        delay_issue.write(i + 1, 0, record['项目名称'], style)
        delay_issue.write(i + 1, 1, record['issue标题'], style)
        delay_issue.write(i + 1, 2, record['逾期天数'], style)
        delay_issue.write(i + 1, 3, record['处理人'], style)
        delay_issue.write(i + 1, 4, record['链接'], style)
        assignee_valid = False
        for k in range(len(assignee_list)):
            if record['处理人'] == assignee_list[k]['处理人']:
                assignee_list[k]['逾期数量'] += 1
                assignee_list[k]['累计逾期天数'] += record['逾期天数']
                assignee_list[k]['平均逾期天数'] = int(assignee_list[k]['累计逾期天数'] / assignee_list[k]['逾期数量'])
                assignee_valid = True
                break
        if not assignee_valid:
            assignee_dict = {'处理人': record['处理人'], '逾期数量': 1, '累计逾期天数': record['逾期天数'], '平均逾期天数': record['逾期天数']}
            assignee_list.append(assignee_dict)
    issue_owner.col(0).width = 7*256
    issue_owner.col(1).width = 12*256
    issue_owner.col(2).width = 12*256
    issue_owner.col(3).width = 12*256
    row0 = ['处理人', '累计逾期数量', '累计逾期天数', '平均逾期天数']
    for i in range(0, len(row0)):
        issue_owner.write(0, i, row0[i], style)
    assignee_list_len = len(assignee_list)
    for i in range(assignee_list_len):
        assignee_list_save = find_delay_max(assignee_list)
        issue_owner.write(i + 1, 0, assignee_list_save['处理人'], style)
        issue_owner.write(i + 1, 1, assignee_list_save['逾期数量'], style)
        issue_owner.write(i + 1, 2, assignee_list_save['累计逾期天数'], style)
        issue_owner.write(i + 1, 3, assignee_list_save['平均逾期天数'], style)
    nddl_issue = workbook.add_sheet('issue统计(无截止日期)')
    nddl_issue.col(0).width = 20*256
    nddl_issue.col(1).width = 120*256
    nddl_issue.col(2).width = 7*256
    nddl_issue.col(3).width = 60*256
    row0 = ['项目名称', 'issue标题', '创建人', '链接']
    for i in range(0, len(row0)):
        nddl_issue.write(0, i, row0[i], style)
    author_list = []
    for i in range(len(no_ddl_data)):
        record = no_ddl_data[i]
        nddl_issue.write(i + 1, 0, record['项目名称'], style)
        nddl_issue.write(i + 1, 1, record['issue标题'], style)
        nddl_issue.write(i + 1, 2, record['创建人'], style)
        nddl_issue.write(i + 1, 3, record['链接'], style)
        author_valid = False
        for k in range(len(author_list)):
            if record['创建人'] == author_list[k]['创建人']:
                author_list[k]['异常数量'] += 1
                author_valid = True
                break
        if not author_valid:
            author_dict = {'创建人': record['创建人'], '异常数量': 1}
            author_list.append(author_dict)
    issue_owner.col(5).width = 7*256
    issue_owner.col(6).width = 12*256
    row0 = ['创建人', '异常数量']
    for i in range(len(row0)):
        issue_owner.write(0, i + 5, row0[i], style)
    author_list_len = len(author_list)
    for i in range(author_list_len):
        author_list_save = find_nddl_max(author_list)
        issue_owner.write(i + 1, 5, author_list_save['创建人'], style)
        issue_owner.write(i + 1, 6, author_list_save['异常数量'], style)
    workbook.save(excelPath)


def get_delay_issue():
    delay_issue_list = []
    no_ddl_list = []
    gitlab_projects = getAllProjects()
    for gitlab_project in gitlab_projects:
        prj_name = gitlab_project.name
        prj_issue_num = len(gitlab_project.issues.list(all=True))
        for i in range(prj_issue_num):
            issue = gitlab_project.issues.get(i + 1)
            issue_ddl = issue.due_date
            issue_state = issue.state
            if issue_ddl is None and issue_state == "opened":
                issue_dict = {'项目名称': prj_name, 'issue标题': issue.title,
                              '创建人': issue.author['name'], '链接': issue.web_url}
                no_ddl_list.append(issue_dict)
                print(issue_dict)
            issue_close_time = issue.closed_at
            issue_due_time = issue.due_date
            if issue_close_time is None and issue_due_time is not None:
                cur_day = datetime.datetime.now().strftime('%Y-%m-%d')
                delay_time = datetime.datetime.strptime(cur_day, '%Y-%m-%d') - \
                             datetime.datetime.strptime(issue_due_time, '%Y-%m-%d')
                delay_days = delay_time.days
                if delay_days > 0:
                    issue_dict = {'项目名称': prj_name, 'issue标题': issue.title, '逾期天数': delay_days,
                                  '处理人': issue.assignees[0]['name'], '链接': issue.web_url}
                    print(issue_dict)
                    delay_issue_list.append(issue_dict)
    return delay_issue_list, no_ddl_list


if __name__ == '__main__':
    delay_info_data, no_ddl_info = get_delay_issue()
    excel_addr = 'D:/' + 'gitlab issue统计' + '_' + datetime.datetime.now().strftime('%Y-%m-%d') + '.xls'
    writeExcel(excel_addr, delay_info_data, no_ddl_info)
