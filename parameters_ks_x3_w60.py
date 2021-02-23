# writing Yealink parameters to an excel spreadsheet
import xlsxwriter

workbook = xlsxwriter.Workbook('parameters.xlsx')

worksheet = workbook.add_worksheet()

var_voicemail_parameter_group = input("Enter your parameter group: ")
var_voicemail_user_name = int(input("Enter the voicemail user: "))
var_sales_floor_3_name = int(input("Enter the Sales Floor 3 user: "))
hs_line_1 = 1
hs_line_2 = 2

worksheet.write('A1', 'group_name')
worksheet.write('B1', 'action')
worksheet.write('C1', 'parameter')
worksheet.write('D1', 'value')
worksheet.write('E1', 'lock')

worksheet.write('A2', var_voicemail_parameter_group)
worksheet.write('B2', 'SET')
worksheet.write('C2', 'account.1.auth_name')
worksheet.write('D2', var_voicemail_user_name)
worksheet.write('E2', 'disabled')

worksheet.write('A3', var_voicemail_parameter_group)
worksheet.write('B3', 'SET')
worksheet.write('C3', 'account.1.user_name')
worksheet.write('D3', var_voicemail_user_name)
worksheet.write('E3', 'disabled')

worksheet.write('A4', var_voicemail_parameter_group)
worksheet.write('B4', 'SET')
worksheet.write('C4', 'account.1.display_mwi.enable')
worksheet.write('D4', 1)
worksheet.write('E4', 'disabled')

worksheet.write('A5', var_voicemail_parameter_group)
worksheet.write('B5', 'SET')
worksheet.write('C5', 'account.1.outbound_proxy.1.address')
worksheet.write('D5', 'proxy-ucc.genband.com')
worksheet.write('E5', 'disabled')

worksheet.write('A6', var_voicemail_parameter_group)
worksheet.write('B6', 'SET')
worksheet.write('C6', 'account.1.outbound_proxy_enable')
worksheet.write('D6', 1)
worksheet.write('E6', 'disabled')

worksheet.write('A7', var_voicemail_parameter_group)
worksheet.write('B7', 'SET')
worksheet.write('C7', 'account.1.password')
worksheet.write('D7', '<password>')
worksheet.write('E7', 'disabled')

worksheet.write('A8', var_voicemail_parameter_group)
worksheet.write('B8', 'SET')
worksheet.write('C8', 'account.1.sip_server.1.address')
worksheet.write('D8', 'att-ni-ics.com')
worksheet.write('E8', 'disabled')

worksheet.write('A9', var_voicemail_parameter_group)
worksheet.write('B9', 'SET')
worksheet.write('C9', 'phone_setting.mail_power_led_flash_enable')
worksheet.write('D9', 1)
worksheet.write('E9', 'disabled')

worksheet.write('A10', var_voicemail_parameter_group)
worksheet.write('B10', 'SET')
worksheet.write('C10', 'account.2.auth_name')
worksheet.write('D10', var_sales_floor_3_name)
worksheet.write('E10', 'disabled')

worksheet.write('A11', var_voicemail_parameter_group)
worksheet.write('B11', 'SET')
worksheet.write('C11', 'account.2.user_name')
worksheet.write('D11', var_sales_floor_3_name)
worksheet.write('E11', 'disabled')

worksheet.write('A12', var_voicemail_parameter_group)
worksheet.write('B12', 'SET')
worksheet.write('C12', 'account.2.display_mwi.enable')
worksheet.write('D12', 1)
worksheet.write('E12', 'disabled')

worksheet.write('A13', var_voicemail_parameter_group)
worksheet.write('B13', 'SET')
worksheet.write('C13', 'account.2.outbound_proxy.1.address')
worksheet.write('D13', 'proxy-ucc.genband.com')
worksheet.write('E13', 'disabled')

worksheet.write('A14', var_voicemail_parameter_group)
worksheet.write('B14', 'SET')
worksheet.write('C14', 'account.2.outbound_proxy_enable')
worksheet.write('D14', 1)
worksheet.write('E14', 'disabled')

worksheet.write('A15', var_voicemail_parameter_group)
worksheet.write('B15', 'SET')
worksheet.write('C15', 'account.2.password')
worksheet.write('D15', '<password>')
worksheet.write('E15', 'disabled')

worksheet.write('A16', var_voicemail_parameter_group)
worksheet.write('B16', 'SET')
worksheet.write('C16', 'account.2.sip_server.1.address')
worksheet.write('D16', 'att-ni-ics.com')
worksheet.write('E16', 'disabled')

worksheet.write('A17', var_voicemail_parameter_group)
worksheet.write('B17', 'SET')
worksheet.write('C17', 'account.2.label')
worksheet.write('D17', var_sales_floor_3_name)
worksheet.write('E17', 'disabled')

worksheet.write('A18', var_voicemail_parameter_group)
worksheet.write('B18', 'SET')
worksheet.write('C18', 'account.2.display_name')
worksheet.write('D18', var_sales_floor_3_name)
worksheet.write('E18', 'disabled')

worksheet.write('A19', var_voicemail_parameter_group)
worksheet.write('B19', 'SET')
worksheet.write('C19', 'handset.1.name')
worksheet.write('D19', 'Backroom')
worksheet.write('E19', 'disabled')

worksheet.write('A20', var_voicemail_parameter_group)
worksheet.write('B20', 'SET')
worksheet.write('C20', 'handset.1.incoming_lines')
worksheet.write('D20', '2, 1')
worksheet.write('E20', 'disabled')

worksheet.write('A21', var_voicemail_parameter_group)
worksheet.write('B21', 'SET')
worksheet.write('C21', 'handset.1.dial_out_lines')
worksheet.write('D21', '2, 1')
worksheet.write('E21', 'disabled')

worksheet.write('A22', var_voicemail_parameter_group)
worksheet.write('B22', 'SET')
worksheet.write('C22', 'handset.1.dial_out_default_line')
worksheet.write('D22', 1)
worksheet.write('E22', 'disabled')

workbook.close()