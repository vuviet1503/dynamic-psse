


import os, sys
# import psse35
import pandas as pd



def input_grid(file):
    # read grid excel file
    sheet_names = ["BUS", "MACHINE", "LOAD", "SHUNT", "LINE", "2 WINDING"]
    xls = pd.read_excel(file, sheet_name=sheet_names, skiprows=1)
    
    bus = xls["BUS"].values.tolist()
    machine = xls["MACHINE"].values.tolist()
    load = xls["LOAD"].values.tolist()
    shunt = xls["SHUNT"].values.tolist()
    line = xls["LINE"].values.tolist()
    twinding = xls["2 WINDING"].values.tolist()

    import psse35
    import psspy

    ierr = psspy.psseinit()
    ierr = psspy.newcase_2(basemva = 100, basefreq = 50)
    
    # add bus
    for row in bus:
        bus_number, bus_name, base_kv, area, zone, owner, code = row
        ierr = psspy.bus_data_4(ibus = bus_number, inode = 0, intgar =[code, area, zone, owner], realar = [base_kv], name = bus_name)
    
    # add machine, plant
    for row in machine:
        bus_number, vsch, pgen = row
        ierr = psspy.plant_data_4(ibus = int(bus_number), inode = 0, realar = [vsch])
        ierr = psspy.machine_data_4(ibus = int(bus_number), id = "1", realar = [pgen])

    # add load
    for row in load:
        bus_number, pload, qload = row
        ierr = psspy.load_data_4(ibus = int(bus_number), id = "1", realar = [pload, qload])
    
    #add shunt
    for row in shunt:
        bus_number, qshunt = row
        ierr = psspy.shunt_data(ibus = int(bus_number), id = "1", realar = [0, qshunt])

    #add branch
    for row in line:
        ibus, jbus, r, x, b, rate = row
        ierr = psspy.branch_data_3(ibus = int(ibus), jbus = int(jbus), realar = [r, x, b], ratings = [rate])

    #add 2 winding
    for row in twinding:
        ibus, jbus, r, x, g, b, rate = row
        ierr, realaro = psspy.two_winding_data_6(ibus = int(ibus), jbus = int(jbus), realari = [r, x], ratings = [rate])

    ierr = psspy.save("test.sav")


def dynamics():
    import psse35
    import psspy
    import dyntools 

    ierr = psspy.case("test.sav")
    ierr = psspy.fdns([0,0,0,1,1,0,99,0])
    psspy.cong(0)
    psspy.conl(0,1,1,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.conl(0,1,3,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.ordr(1)
    psspy.fact()
    psspy.tysl(0)
    psspy.dyre_new([1,1,1,1],r"""C:\Users\Admin\Documents\PTI\PSSE35\EXAMPLE\savnw.dyr""","","","")
    psspy.bsys(1,0,[ 13.8, 500.],0,[],1,[101],0,[],0,[])
    psspy.bsys(1,0,[ 13.8, 500.],0,[],1,[101],0,[],0,[])
    psspy.chsb(1,0,[1,9,1,1,2,0])
    psspy.strt_2([0,0],"fault.out")
    psspy.run(0,-0.02,999,1,0)
    psspy.run(0, 10.0,999,1,0)
    psspy.dist_3phase_bus_fault(101,0,1, 21.6,[0.0,-0.2E+10])
    psspy.change_channel_out_file("fault.out")
    psspy.run(0, 10.1,999,1,0)
    psspy.dist_clear_fault(1)
    psspy.change_channel_out_file("fault.out")
    psspy.run(0, 20.0,999,1,0)



def output():
    import os
    import psse35
    import dyntools
    import pandas as pd

    # Đường dẫn file .out (cần thay đổi nếu file ở vị trí khác)
    outfile = "fault.out"

    # Kiểm tra file tồn tại
    if not os.path.exists(outfile):
        print(f"⚠️ File '{outfile}' không tồn tại!")
        exit()

    # Sử dụng dyntools để load dữ liệu
    chnfobj = dyntools.CHNF(outfile)

    # Lấy thông tin dữ liệu
    short_title, chanid, chandata = chnfobj.get_data()

    # Hiển thị tiêu đề file .out
    print(f"\n📄 Tiêu đề file: {short_title}\n")

    # Hiển thị danh sách các kênh dữ liệu
    print("🔍 Danh sách các kênh dữ liệu:")
    for ch_id, ch_name in chanid.items():
        print(f"  - Kênh {ch_id}: {ch_name}")

    # Chuyển dữ liệu thành DataFrame
    df = pd.DataFrame(chandata)

    # Đường dẫn file CSV để lưu
    csv_file = r"C:\Users\Admin\Desktop\code\dynamicsPSSe\output_data.csv"

    # Lưu dữ liệu ra CSV
    df.to_csv(csv_file, index=False)
    print(f"\n✅ Dữ liệu đã được lưu vào: {csv_file}")

    import matplotlib.pyplot as plt

    # Chọn kênh cần vẽ (VD: kênh 1)
    channel_number = 1

    # Kiểm tra kênh hợp lệ
    if channel_number not in chanid:
        print(f"⚠️ Kênh {channel_number} không tồn tại!")
    else:
        # Lấy dữ liệu thời gian và giá trị kênh
        time_data = chandata['time']
        channel_data = chandata[channel_number]

        # Vẽ đồ thị
        plt.figure(figsize=(10, 5))
        plt.plot(time_data, channel_data, label=chanid[channel_number], color='b')
        plt.xlabel("Time (s)")
        plt.ylabel("Value")
        plt.title(f"Biểu đồ kênh {channel_number}: {chanid[channel_number]}")
        plt.legend()
        plt.grid(True)
        plt.show()


if __name__ == "__main__":
    file = r"C:\Users\Admin\Desktop\code\dynamicsPSSe\SAVNW.xlsx"
    input_grid(file)
    dynamics()
    output()