import time
import itertools
import requests,openpyxl,datetime,psycopg2
class TCM:
    # -----定义一些常用参数变量--------
    def __init__(self):
        self.excel = openpyxl.load_workbook(r'C:\excel_to_TCMAPI.xlsx')
        self.testplan_sheet = self.excel['test_plan']
        self.testcase_sheet = self.excel['test_case']
        self.testplan_name=self.testplan_sheet['B1'].value
        self.project_name=self.testplan_sheet['B2'].value
        self.testplan_phase=self.testplan_sheet['B3'].value
        self.testcase_sheet_max_rows = self.testcase_sheet.max_row
        self.token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyIiOiJHdWVzdCJ9.JQHky3HpSFnUkTn9TPA95bzANa2-9B3yXYALa36vIp8"
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
            "Authorization": f"Bearer {self.token}"
                       }  # 这个token值是永恒不变的，根据手动登录后自己用户名密码生产的

        print('请等待,正在连接TCM数据库')
    # -----定义一些常用参数变量--------

    def exceldata_switchTo_tcmdata_thenrequest(self):
        while True:
            try:
                db = psycopg2.connect(host="10.30.187.103", user='tcm_prd', password='YzkTfW4n!@@HA&FB', port=5432, database='tcm', connect_timeout=60)
                cursor = db.cursor()
                print('TCM数据库连接成功')
                break
            except Exception as e:
                print('TCM数据库连接失败:' + str(e) + '10秒后自动尝试重新连接')
                time.sleep(10)
                print('正在重新连接TCM数据库')
        for every_row in range(2, self.testcase_sheet_max_rows + 1):# 这里是匹配testplan sheet的最大行数，除去第一行表头，+1是因为range(2,9）行是循环2到8行.所以加1刚刚好
            try:
                every_row_list = []  # 给每一行建立一个空列表来装每行一具体的值
                for cell in self.testcase_sheet[every_row]:  # testplan_sheet[i] i是就代表testplan sheet的第几行，cell是代表这一行中的每一个单元格
                    every_row_list.append(cell.value)
                every_row_list[4] = every_row_list[4].strftime("%Y-%m-%d")  # 开始日期转化为字符串格式，因为从excel读的值datetime.datetime格式，开始时间栏位本身就是字符串，不需要转
                every_row_list[5] = every_row_list[5].strftime("%Y-%m-%d")  # 结束日期转化为字符串格式，因为从excel读的值datetime.datetime格式，结束时间栏位本身就是字符串，不需要转
                #把excel的时间部分转化为TCM需要的格式
                start_time=every_row_list[4]+' '+every_row_list[6]#开始日期和时间以字符串的形式连接起来"2021-01-02 05:30"象这样
                start_time_timeformat = datetime.datetime.strptime(start_time, "%Y-%m-%d %H:%M")#转为时间格式，为了后面转换世界时间准备
                start_word_time= start_time_timeformat-datetime.timedelta(hours=8)#TCM回填的是世界时间，也就是是北京时间减8小时
                start_word_time_str=str(start_word_time)#转为字符串格式
                start_word_time_tcm=start_word_time_str[0:10]+'T'+start_word_time_str[11:]+".000Z"#转为TCM需要的格式2024-03-21T21:07:00.000Z
                end_time = every_row_list[5] + ' ' + every_row_list[7]  # 结束日期和时间以字符串的形式连接起来"2021-01-02 05:30"象这样
                end_time_timeformat = datetime.datetime.strptime(end_time, "%Y-%m-%d %H:%M")  # 转为时间格式，为了后面转换世界时间准备
                end_word_time = end_time_timeformat - datetime.timedelta(hours=8)  # TCM回填的是世界时间，也就是是北京时间减8小时
                end_word_time_str = str(end_word_time)  # 转为字符串格式
                end_word_time_tcm = end_word_time_str[0:10] + 'T' + end_word_time_str[11:] + ".000Z"  # 转为TCM需要的格式2024-03-21T21:07:00.000Z
                # 把excel的时间部分转化为TCM需要的格式

                if every_row_list[11]!=None:
                    unattened=float(every_row_list[11])
                else:
                    unattened=0


                #把excel nonTestDurationDetailsreson_time_comment栏位转为TCM需要的格式
                if every_row_list[8]==None: #代表没有填nontest原因
                    nonTestDuration=0 #回填时间就是0
                    nonTestDurationDetails=[]# 回填空列表
                else:
                    dict_nontestreson = {'Unattend': 1, 'Test case isolation': 2, 'Issue verify': 3, 'RD debug': 4, 'RD rework': 5, 'SUT lending': 6, 'Others': 7}#excel填汉字，实际上传TCM是对应的原因ID，所以先准备好字典
                    nonTestDurationDetails = []
                    nonTestDuration=0
                    nontestresontimecomment_list = every_row_list[8].split('\n')#把每一行的nontest原因时间按换行符转为列表['Unattend_1.5_吃饭', 'Issue verify_2']
                    for every_reason in nontestresontimecomment_list:
                        dict={}#定义一个空字典，等会把列表的值添加到字典里，再append到nonTestDurationDetails列表里，因为TCM 上传的是nonTestDurationDetails列表，列表里加字典
                        nontestresontimecomment_list_every_list = every_reason.split('_')#['Unattend_1.5_吃饭', 'Issue verify_2']把列表里的每个元素按下划线再次转为列表['Unattend','1.5','吃饭']
                        nontestresontimecomment_list_every_list[1] = float(nontestresontimecomment_list_every_list[1])#把1.5转为浮点型, 方便计算总和
                        dict['reasonId'] = dict_nontestreson[nontestresontimecomment_list_every_list[0]]#像字典添加键值对，公式例子:dict[’a‘]=3 即可添加{”a":3},dict_nontestreson[nontestresontimecommont_list_every_list[0]的值是1
                        dict['durationTime'] = nontestresontimecomment_list_every_list[1]#添加时间键值对
                        if len(nontestresontimecomment_list_every_list) == 2:#如果等于2就代没有备注
                            dict['comment'] = None
                        else:
                            dict['comment'] = nontestresontimecomment_list_every_list[2]#添加备注键值对
                        nonTestDurationDetails.append(dict)
                        nonTestDuration=nonTestDuration+nontestresontimecomment_list_every_list[1]
                # 把excel nonTestDurationDetailsreson_time_comment栏位转为TCM需要的格式

                # 根据excel plan name case ID等咨询查询数据库确定json data里caseSteps的内容以及查找testrunid"caseSteps":[{"id":"ebaa187b-333c-47f2-8a31-83ef0c51c8a2","actualOutCome":null,"status":"Pass","testResultComment":null,"issues":[],"attachments":[]}],
                print('case id:'+every_row_list[0]+' && '+'map config:'+every_row_list[2]+' 即将开始执行结果回传')
                print('正在查询接口')
                sql_planid = f"SELECT id FROM test_plan.plan where name='{self.testplan_name}'"  # 需要执行的语句, 根据plan名字查询id，id是唯一值
                cursor.execute(sql_planid)#执行语句
                result_return_planid = cursor.fetchone()#result_return_planid存储返回的结果，因为这里ID是唯一值，所以用fetchone()返回元组，返回多个值用fetchall()返回字典
                plan_id=result_return_planid[0]#返回的元组只有1个值比如(87,)取第一个值就是plan ID
                sql_iteration_id=f"SELECT id FROM test_plan.iteration where name='{every_row_list[1]}' and plan_id='{plan_id}'"#根据plan id加iteration name锁定iteration id
                cursor.execute(sql_iteration_id)
                result_return_iteration_id = cursor.fetchone()
                iteration_id=result_return_iteration_id[0]
                sql_iteration_case_group_id = f"SELECT id FROM test_plan.iteration_case_group where iteration_id='{iteration_id}' and case_id='{every_row_list[0]}'"
                cursor.execute(sql_iteration_case_group_id)
                result_return_iterationcasegroupid = cursor.fetchone()
                if result_return_iterationcasegroupid==None:#没找到case
                    reason=self.testplan_name+' plan中没有拉test case '+every_row_list[0]
                    print(reason+','+'即将开始下一份case回传')
                    self.testcase_sheet.cell(column=13, row=every_row, value="回传失败原因:" + reason)
                    self.excel.save(r'C:\excel_to_TCMAPI.xlsx')
                else: #代表有找到case
                    iteration_case_group_id = result_return_iterationcasegroupid[0]
                    sql_iteration_case_group_child_id = f"SELECT id,test_case_type FROM test_plan.iteration_case_group_child where iteration_case_group_id='{iteration_case_group_id}'"
                    cursor.execute(sql_iteration_case_group_child_id)
                    result_return_iterationcasegroupchlidid = cursor.fetchall()#这里因为返回child case，可能是多个结果，所以用fetchall,例如原始case带出internal 这种，返回列表[('3f8ff388-f915-4a12-aaa3-7f662a63cb07', ), ('45c62a69-ed88-46fd-ac35-71bd31b1bc03', ), ('5f0aeb50-1515-47f2-bb40-662bd1b7cf2c', )]
                    if len(result_return_iterationcasegroupchlidid)==1:# 等与1就是没有带出internal case
                        iteration_case_group_child_id = []  # 为把返回的列表，但列表里是元组，直接全部转列表做准备
                        iteration_case_group_child_id.append(result_return_iterationcasegroupchlidid[0][0])
                        sql_project_id = f"SELECT qt_project_id FROM qt_project_manage.qt_project where project_name='{self.project_name}'"
                        cursor.execute(sql_project_id)
                        return_result_projectid = cursor.fetchone()
                        project_id = return_result_projectid[0]
                        #sql_qt_dut_phase_management_id = f"SELECT id,configuration_name FROM qt_project_manage.qt_dut_phase_management where qt_project_id={project_id} and phase='{self.testplan_phase}'"
                        sql_qt_dut_phase_management_id = f"SELECT qdpm.id, qpdc.configuration_name FROM qt_project_manage.qt_dut_phase_management qdpm JOIN qt_project_manage.qt_phase_detail_config qpdc ON qdpm.qt_phase_detail_config_id =qpdc.id JOIN qt_project_manage.qt_phase_detail qpd ON qpdc.qt_phase_detail_id =qpd.id JOIN qt_project_manage.qt_dut_phase qdp ON qpd.qt_dut_phase_id =qdp.id WHERE qdp.custom_phase = '{self.testplan_phase}' AND qdp.qt_project_id = {project_id}"
                        cursor.execute(sql_qt_dut_phase_management_id)
                        return_result_qtdutphasemanagement = cursor.fetchall()
                        configs = return_result_qtdutphasemanagement
                        SUT_dict = {}  # 最终变这样{'SIT SKU3': '3559158e-58e6-4b73-ace5-15502f42a6d9', 'SIT SKU2': '9010e4e6-3dfb-4c29-9c4b-0134cde4e0cc'}
                        for each_config in configs:
                            SUT_dict[each_config[1]] = each_config[0]
                            sql_sub_config = f"SELECT id,configuration_name FROM qt_project_manage.qt_dut_phase_management_sub_config where superior_config_id='{each_config[0]}'"
                            cursor.execute(sql_sub_config)
                            sub_config = cursor.fetchall()
                            if sub_config != None:
                                for each_sub_config in sub_config:
                                    SUT_dict[each_sub_config[1]] = each_sub_config[0]

                        sql_iteration_config_id = f"SELECT id FROM test_plan.iteration_config where qt_dut_phase_management_id='{SUT_dict[every_row_list[2]]}' and iteration_id='{iteration_id}'"
                        cursor.execute(sql_iteration_config_id)
                        return_result_iterationconfifid = cursor.fetchone()
                        if return_result_iterationconfifid==None:
                            sql_iteration_config_id = f"SELECT id FROM test_plan.iteration_config where qt_dut_phase_management_sub_config_id='{SUT_dict[every_row_list[2]]}' and iteration_id='{iteration_id}'"
                            cursor.execute(sql_iteration_config_id)
                            return_result_iterationconfifid = cursor.fetchone()
                        iteration_config_id = return_result_iterationconfifid[0]
                        case_group_map_id_list = []
                        sql_case_group_map_id = f"SELECT id,iteration_config_id FROM test_plan.case_group_map where iteration_case_group_id='{iteration_case_group_id}' and iteration_config_id='{iteration_config_id}'"
                        cursor.execute(sql_case_group_map_id)
                        result_return_casegroupmapid = cursor.fetchone()  # 因为指定了config说以这里只会返回一个id
                        case_group_map_id = result_return_casegroupmapid[0]
                        case_group_map_id_list.append(case_group_map_id)
                        combinations = list(itertools.product(case_group_map_id_list, iteration_case_group_child_id))  # 这里是列表组合，把map id列表排列组合iteration_case_group_child_id,[('cc19bdfb-460e-469d-a609-e60128381e5d', '45c62a69-ed88-46fd-ac35-71bd31b1bc03'), ('cc19bdfb-460e-469d-a609-e60128381e5d', '5f0aeb50-1515-47f2-bb40-662bd1b7cf2c'), ('abb5879e-d0c8-4901-bf6e-2a628a0631b8', '45c62a69-ed88-46fd-ac35-71bd31b1bc03'), ('abb5879e-d0c8-4901-bf6e-2a628a0631b8', '5f0aeb50-1515-47f2-bb40-662bd1b7cf2c'), ('f8039461-ac5e-4610-af01-9136c911cd66', '45c62a69-ed88-46fd-ac35-71bd31b1bc03'), ('f8039461-ac5e-4610-af01-9136c911cd66', '5f0aeb50-1515-47f2-bb40-662bd1b7cf2c')]
                        case_group_child_map_id = []
                        for combo in combinations:  # 比如mapid=[1,2,3],iteration_case_group_child_id=[3,4],combo就是元组(1,3),(1,4),(2,3),(2,4),(3,3),(3,4)
                            sql_case_group_child_map_id = f"SELECT id FROM test_plan.case_group_child_map where case_group_map_id='{combo[0]}' and iteration_case_group_child_id='{combo[1]}'"  # 根据map id 和 iteration_case_group_child_id锁定case_group_child_map_id
                            cursor.execute(sql_case_group_child_map_id)
                            return_result_case_group_child_map_id = cursor.fetchone()  # 一个map id一个iteration_case_group_child_id结合查询只会返回一个case_group_child_map_id
                            case_group_child_map_id.append(return_result_case_group_child_map_id[0])
                        case_group_child_cycle_id = []  # 这里一个map id 可能带出多个cycle id, 如果一个case 增加了cycle，就会在这里查询出现多个cycle id
                        for every_case_group_child_map_id in case_group_child_map_id:
                            sql_case_group_child_cycle_id = f"SELECT id FROM test_plan.case_group_child_cycle where case_group_child_map_id='{every_case_group_child_map_id}'"
                            cursor.execute(sql_case_group_child_cycle_id)
                            return_result_case_group_child_cycle_id = cursor.fetchall()  # 所以这里用all
                            for every_child_cycyle_id in return_result_case_group_child_cycle_id:
                                case_group_child_cycle_id.append(every_child_cycyle_id[0])
                        testrun_id = []  # 再根据case_group_child_cycle_id锁定testrun_id
                        for every_case_group_child_cycle_id in case_group_child_cycle_id:
                            sql_testrun_id = f"SELECT id FROM test_plan.test_results where case_group_child_cycle_id='{every_case_group_child_cycle_id}'"
                            cursor.execute(sql_testrun_id)
                            return_result_testrun_id = cursor.fetchone()
                            testrun_id.append(return_result_testrun_id[0])
                        ###截至到这里这份case的对应的config的所有test run id就找到了.接下来根据test run id找case steps
                        print(f'接口查询成功,此case对应TCM需要回填的case数量为{len(testrun_id)}份')
                        for every_testrunid in testrun_id:
                            number = testrun_id.index(every_testrunid) + 1
                            print(f'正在回传第{number}份')
                            caseSteps = []
                            issues = []
                            if every_row_list[10] == None:  # 代表没有issue,pass
                                sql_casesteps = f"SELECT id FROM test_plan.test_result_steps where test_result_id={every_testrunid}"
                                cursor.execute(sql_casesteps)
                                case_steps = cursor.fetchall()  # 如果是5步，[('5276810c-d8a3-4f22-ba36-692fc9169344',), ('0904a790-9268-4dd4-b226-a6bcc41cbbce',), ('9e572f3d-5c34-4568-9208-d7141fc0b54f',), ('a67a5904-eeaa-4896-b9ae-578ca0350b55',), ('daeab79c-37a5-4290-8802-e345c5a17949',)]
                                for every_step in case_steps:  # 构建每一步的字典
                                    every_step_dict = {}
                                    every_step_dict['id'] = every_step[0]
                                    every_step_dict['actualOutCome'] = None
                                    every_step_dict['status'] = 'Pass'
                                    every_step_dict['testResultComment'] = None
                                    every_step_dict['issues'] = []
                                    every_step_dict['attachments'] = []
                                    caseSteps.append(every_step_dict)
                            else:  # 就代表有fail
                                sql_casesteps = f"SELECT id FROM test_plan.test_result_steps where test_result_id={every_testrunid}"
                                cursor.execute(sql_casesteps)
                                case_steps = cursor.fetchall()  # 如果是5步，[('5276810c-d8a3-4f22-ba36-692fc9169344',), ('0904a790-9268-4dd4-b226-a6bcc41cbbce',), ('9e572f3d-5c34-4568-9208-d7141fc0b54f',), ('a67a5904-eeaa-4896-b9ae-578ca0350b55',), ('daeab79c-37a5-4290-8802-e345c5a17949',)]
                                issue_list = every_row_list[10].split('$')
                                ##接下来做第一步打pass其余步骤fail
                                if len(case_steps) == 1:  # 代表本身就只有1步：
                                    every_step_dict = {}
                                    for every_issue in issue_list:
                                        issue_dict = {}
                                        issue_dict['issueId'] = every_issue[0:6]
                                        issue_dict['issueTitle'] = every_issue[6:]
                                        issues.append(issue_dict)
                                    first_step_id = case_steps[0][0]  # 第一步的id
                                    every_step_dict['id'] = first_step_id
                                    every_step_dict['actualOutCome'] = None
                                    every_step_dict['status'] = 'Fail'
                                    every_step_dict['testResultComment'] = None
                                    every_step_dict['issues'] = issues
                                    every_step_dict['attachments'] = []
                                    caseSteps.append(every_step_dict)
                                else:  # 就是有多步骤, 把第一步打fail其余pass
                                    first_step_dict = {}
                                    for every_issue in issue_list:
                                        issue_dict = {}
                                        issue_dict['issueId'] = every_issue[0:6]
                                        issue_dict['issueTitle'] = every_issue[6:]
                                        issues.append(issue_dict)
                                    first_step_id = case_steps[0][0]  # 第一步的id
                                    first_step_dict['id'] = first_step_id
                                    first_step_dict['actualOutCome'] = None
                                    first_step_dict['status'] = 'Fail'
                                    first_step_dict['testResultComment'] = None
                                    first_step_dict['issues'] = issues
                                    first_step_dict['attachments'] = []
                                    caseSteps.append(first_step_dict)
                                    case_steps.pop(0)  # 第一步打fail后删除，从第二步开始都pass
                                    for every_step in case_steps:  # 构建每一步的字典
                                        every_step_dict = {}
                                        every_step_dict['id'] = every_step[0]
                                        every_step_dict['actualOutCome'] = None
                                        every_step_dict['status'] = 'Pass'
                                        every_step_dict['testResultComment'] = None
                                        every_step_dict['issues'] = []
                                        every_step_dict['attachments'] = []
                                        caseSteps.append(every_step_dict)
                            data = {
                                "customerRunId": None,
                                "hyperlink": None,
                                "actualConfig": every_row_list[2],
                                "actualDut": every_row_list[3],
                                "actualStartTime": start_word_time_tcm,
                                "actualEndTime": end_word_time_tcm,
                                "nonTestDuration": nonTestDuration,
                                "caseSteps": caseSteps,
                                "unattened": unattened,
                                "nonTestDurationDetails": nonTestDurationDetails
                            }  # put传到TCM的json数据
                            code_list = []
                            url = f"https://tms.wistron.com/qt/api/v2/testplans/test_run/{every_testrunid}?complete=true"
                            re = requests.put(url=url, headers=self.headers, json=data)
                            code_list.append(re.status_code)
                            if re.status_code == 200:
                                print(f'第{number}份回传成功')
                            else:
                                print(f'第{number}份回传失败')
                        if 200 in code_list and len(code_list) == 1:  #
                            self.testcase_sheet.cell(column=13, row=every_row, value='Success')
                            self.excel.save(r'C:\excel_to_TCMAPI.xlsx')
                        else:
                            self.testcase_sheet.cell(column=13, row=every_row, value='Fail')
                            self.excel.save(r'C:\excel_to_TCMAPI.xlsx')

                    else:  #就是有带出internal case
                        iteration_case_group_child_id = []#为把返回的列表，但列表里是元组，直接全部转列表做准备
                        for every_child_case in result_return_iterationcasegroupchlidid:
                            if every_child_case[1] !='Original': #排除origina的child case
                                iteration_case_group_child_id.append(every_child_case[0])###变为['45c62a69-ed88-46fd-ac35-71bd31b1bc03', '5f0aeb50-1515-47f2-bb40-662bd1b7cf2c']，，，every_child_case[0]就是代表（’jjr',)取第一个元素jjr
                        sql_project_id = f"SELECT qt_project_id FROM qt_project_manage.qt_project where project_name='{self.project_name}'"
                        cursor.execute(sql_project_id)
                        return_result_projectid = cursor.fetchone()
                        project_id = return_result_projectid[0]
                        #sql_qt_dut_phase_management_id = f"SELECT id,configuration_name FROM qt_project_manage.qt_dut_phase_management where qt_project_id={project_id} and phase='{self.testplan_phase}'"
                        sql_qt_dut_phase_management_id = f"SELECT qdpm.id, qpdc.configuration_name FROM qt_project_manage.qt_dut_phase_management qdpm JOIN qt_project_manage.qt_phase_detail_config qpdc ON qdpm.qt_phase_detail_config_id =qpdc.id JOIN qt_project_manage.qt_phase_detail qpd ON qpdc.qt_phase_detail_id =qpd.id JOIN qt_project_manage.qt_dut_phase qdp ON qpd.qt_dut_phase_id =qdp.id WHERE qdp.custom_phase = '{self.testplan_phase}' AND qdp.qt_project_id = {project_id}"
                        cursor.execute(sql_qt_dut_phase_management_id)
                        return_result_qtdutphasemanagement = cursor.fetchall()
                        configs = return_result_qtdutphasemanagement
                        SUT_dict = {}#最终变这样{'SIT SKU3': '3559158e-58e6-4b73-ace5-15502f42a6d9', 'SIT SKU2': '9010e4e6-3dfb-4c29-9c4b-0134cde4e0cc'}
                        for each_config in configs:
                            SUT_dict[each_config[1]] = each_config[0]
                            sql_sub_config = f"SELECT id,configuration_name FROM qt_project_manage.qt_dut_phase_management_sub_config where superior_config_id='{each_config[0]}'"
                            cursor.execute(sql_sub_config)
                            sub_config = cursor.fetchall()
                            if sub_config != None:
                                for each_sub_config in sub_config:
                                    SUT_dict[each_sub_config[1]] = each_sub_config[0]
                        sql_iteration_config_id = f"SELECT id FROM test_plan.iteration_config where qt_dut_phase_management_id='{SUT_dict[every_row_list[2]]}' and iteration_id='{iteration_id}'"
                        cursor.execute(sql_iteration_config_id)
                        return_result_iterationconfifid = cursor.fetchone()
                        if return_result_iterationconfifid == None:
                            sql_iteration_config_id = f"SELECT id FROM test_plan.iteration_config where qt_dut_phase_management_sub_config_id='{SUT_dict[every_row_list[2]]}' and iteration_id='{iteration_id}'"
                            cursor.execute(sql_iteration_config_id)
                            return_result_iterationconfifid = cursor.fetchone()
                        iteration_config_id = return_result_iterationconfifid[0]
                        case_group_map_id_list=[]
                        sql_case_group_map_id = f"SELECT id,iteration_config_id FROM test_plan.case_group_map where iteration_case_group_id='{iteration_case_group_id}' and iteration_config_id='{iteration_config_id}'"
                        cursor.execute(sql_case_group_map_id)
                        result_return_casegroupmapid = cursor.fetchone()#因为指定了config说以这里只会返回一个id
                        case_group_map_id =result_return_casegroupmapid[0]
                        case_group_map_id_list.append(case_group_map_id)
                        combinations = list(itertools.product(case_group_map_id_list, iteration_case_group_child_id))#这里是列表组合，把map id列表排列组合iteration_case_group_child_id,[('cc19bdfb-460e-469d-a609-e60128381e5d', '45c62a69-ed88-46fd-ac35-71bd31b1bc03'), ('cc19bdfb-460e-469d-a609-e60128381e5d', '5f0aeb50-1515-47f2-bb40-662bd1b7cf2c'), ('abb5879e-d0c8-4901-bf6e-2a628a0631b8', '45c62a69-ed88-46fd-ac35-71bd31b1bc03'), ('abb5879e-d0c8-4901-bf6e-2a628a0631b8', '5f0aeb50-1515-47f2-bb40-662bd1b7cf2c'), ('f8039461-ac5e-4610-af01-9136c911cd66', '45c62a69-ed88-46fd-ac35-71bd31b1bc03'), ('f8039461-ac5e-4610-af01-9136c911cd66', '5f0aeb50-1515-47f2-bb40-662bd1b7cf2c')]
                        case_group_child_map_id = []
                        for combo in combinations: #比如mapid=[1,2,3],iteration_case_group_child_id=[3,4],combo就是元组(1,3),(1,4),(2,3),(2,4),(3,3),(3,4)
                            sql_case_group_child_map_id = f"SELECT id FROM test_plan.case_group_child_map where case_group_map_id='{combo[0]}' and iteration_case_group_child_id='{combo[1]}'"#根据map id 和 iteration_case_group_child_id锁定case_group_child_map_id
                            cursor.execute(sql_case_group_child_map_id)
                            return_result_case_group_child_map_id = cursor.fetchone()# 一个map id一个iteration_case_group_child_id结合查询只会返回一个case_group_child_map_id
                            case_group_child_map_id.append(return_result_case_group_child_map_id[0])
                        case_group_child_cycle_id = []#这里一个map id 可能带出多个cycle id, 如果一个case 增加了cycle，就会在这里查询出现多个cycle id
                        for every_case_group_child_map_id in case_group_child_map_id:
                            sql_case_group_child_cycle_id = f"SELECT id FROM test_plan.case_group_child_cycle where case_group_child_map_id='{every_case_group_child_map_id}'"
                            cursor.execute(sql_case_group_child_cycle_id)
                            return_result_case_group_child_cycle_id = cursor.fetchall()#所以这里用all
                            for every_child_cycyle_id in return_result_case_group_child_cycle_id:
                                case_group_child_cycle_id.append(every_child_cycyle_id[0])
                        testrun_id = []# 再根据case_group_child_cycle_id锁定testrun_id
                        for every_case_group_child_cycle_id in case_group_child_cycle_id:
                            sql_testrun_id = f"SELECT id FROM test_plan.test_results where case_group_child_cycle_id='{every_case_group_child_cycle_id}'"
                            cursor.execute(sql_testrun_id)
                            return_result_testrun_id = cursor.fetchone()
                            testrun_id.append(return_result_testrun_id[0])
                       ###截至到这里这份case的对应的config的所有test run id就找到了.接下来根据test run id找case steps
                        print(f'接口查询成功,此case对应TCM需要回填的case数量为{len(testrun_id)}份')
                        for every_testrunid in testrun_id:
                            number=testrun_id.index(every_testrunid)+1
                            print(f'正在回传第{number}份')
                            caseSteps = []
                            issues=[]
                            if every_row_list[10]==None:#代表没有issue,pass
                                sql_casesteps=f"SELECT id FROM test_plan.test_result_steps where test_result_id={every_testrunid}"
                                cursor.execute(sql_casesteps)
                                case_steps=cursor.fetchall()#如果是5步，[('5276810c-d8a3-4f22-ba36-692fc9169344',), ('0904a790-9268-4dd4-b226-a6bcc41cbbce',), ('9e572f3d-5c34-4568-9208-d7141fc0b54f',), ('a67a5904-eeaa-4896-b9ae-578ca0350b55',), ('daeab79c-37a5-4290-8802-e345c5a17949',)]
                                for every_step in  case_steps:#构建每一步的字典
                                    every_step_dict={}
                                    every_step_dict['id']=every_step[0]
                                    every_step_dict['actualOutCome']=None
                                    every_step_dict['status']='Pass'
                                    every_step_dict['testResultComment']=None
                                    every_step_dict['issues']=[]
                                    every_step_dict['attachments']=[]
                                    caseSteps.append(every_step_dict)
                            else:#就代表有fail
                                sql_casesteps = f"SELECT id FROM test_plan.test_result_steps where test_result_id={every_testrunid}"
                                cursor.execute(sql_casesteps)
                                case_steps = cursor.fetchall()  # 如果是5步，[('5276810c-d8a3-4f22-ba36-692fc9169344',), ('0904a790-9268-4dd4-b226-a6bcc41cbbce',), ('9e572f3d-5c34-4568-9208-d7141fc0b54f',), ('a67a5904-eeaa-4896-b9ae-578ca0350b55',), ('daeab79c-37a5-4290-8802-e345c5a17949',)]
                                issue_list=every_row_list[10].split('$')
                                ##接下来做第一步打pass其余步骤fail
                                if len(case_steps)==1:#代表本身就只有1步：
                                    every_step_dict = {}
                                    for every_issue in issue_list:
                                        issue_dict = {}
                                        issue_dict['issueId']=every_issue[0:6]
                                        issue_dict['issueTitle']=every_issue[6:]
                                        issues.append(issue_dict)
                                    first_step_id=case_steps[0][0]#第一步的id
                                    every_step_dict['id']=first_step_id
                                    every_step_dict['actualOutCome']=None
                                    every_step_dict['status']='Fail'
                                    every_step_dict['testResultComment']=None
                                    every_step_dict['issues']=issues
                                    every_step_dict['attachments']=[]
                                    caseSteps.append(every_step_dict)
                                else:#就是有多步骤, 把第一步打fail其余pass
                                    first_step_dict = {}
                                    for every_issue in issue_list:
                                        issue_dict = {}
                                        issue_dict['issueId'] = every_issue[0:6]
                                        issue_dict['issueTitle'] = every_issue[6:]
                                        issues.append(issue_dict)
                                    first_step_id = case_steps[0][0]  # 第一步的id
                                    first_step_dict['id'] = first_step_id
                                    first_step_dict['actualOutCome'] = None
                                    first_step_dict['status'] = 'Fail'
                                    first_step_dict['testResultComment'] = None
                                    first_step_dict['issues'] = issues
                                    first_step_dict['attachments'] = []
                                    caseSteps.append(first_step_dict)
                                    case_steps.pop(0)#第一步打fail后删除，从第二步开始都pass
                                    for every_step in case_steps:  # 构建每一步的字典
                                        every_step_dict = {}
                                        every_step_dict['id'] = every_step[0]
                                        every_step_dict['actualOutCome'] = None
                                        every_step_dict['status'] = 'Pass'
                                        every_step_dict['testResultComment'] = None
                                        every_step_dict['issues'] = []
                                        every_step_dict['attachments'] = []
                                        caseSteps.append(every_step_dict)
                            data={
                                "customerRunId": None,
                                "hyperlink": None,
                                "actualConfig": every_row_list[2],
                                "actualDut": every_row_list[3],
                                "actualStartTime": start_word_time_tcm,
                                "actualEndTime": end_word_time_tcm,
                                "nonTestDuration": nonTestDuration,
                                "caseSteps": caseSteps,
                                "unattened": unattened,
                                "nonTestDurationDetails": nonTestDurationDetails
                                 } #put传到TCM的json数据
                            code_list = []
                            url=f"https://tms.wistron.com/qt/api/v2/testplans/test_run/{every_testrunid}?complete=true"
                            re=requests.put(url=url,headers=self.headers,json=data)
                            code_list.append(re.status_code)
                            if re.status_code==200:
                                print(f'第{number}份回传成功')
                            else:
                                print(f'第{number}份回传失败')

                        if  200 in code_list and len(code_list)==1:#
                            self.testcase_sheet.cell(column=13, row=every_row, value='Success')
                            self.excel.save(r'C:\excel_to_TCMAPI.xlsx')
                        else:
                            self.testcase_sheet.cell(column=13, row=every_row, value='Fail')
                            self.excel.save(r'C:\excel_to_TCMAPI.xlsx')
            except Exception as e:
                print("回传失败:"+str(e))
                self.testcase_sheet.cell(column=13, row=every_row, value="回传失败原因:"+str(e))
                self.excel.save(r'C:\excel_to_TCMAPI.xlsx')
                continue
        cursor.close()
        db.close()

if __name__ == "__main__":
        case_go = TCM()
        case_go.exceldata_switchTo_tcmdata_thenrequest()



































