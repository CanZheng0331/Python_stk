import time
from tqdm import tqdm
startTime = time.time()
from comtypes.gen import STKObjects, STKUtil, AgStkGatorLib
from comtypes.client import CreateObject, GetActiveObject, GetEvents, CoGetObject, ShowEvents
from ctypes import *
import comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0
from comtypes import GUID
from comtypes import helpstring
from comtypes import COMMETHOD
from comtypes import dispid
from ctypes.wintypes import VARIANT_BOOL
from ctypes import HRESULT
from comtypes import BSTR
from comtypes.automation import VARIANT
from comtypes.automation import _midlSAFEARRAY
from comtypes import CoClass
from comtypes import IUnknown
import comtypes.gen._00DD7BD4_53D5_4870_996B_8ADB8AF904FA_0_1_0
import comtypes.gen._8B49F426_4BF0_49F7_A59B_93961D83CB5D_0_1_0
from comtypes.automation import IDispatch
import comtypes.gen._42D2781B_8A06_4DB2_9969_72D6ABF01A72_0_1_0
from comtypes import DISPMETHOD, DISPPROPERTY, helpstring
import random
import csv

#from Init_Sat import Create_satellite
#from Init_User import CreateUsers
#from CheckAccess import ComputeAndCheckAccess

"""
SET TO TRUE TO USE ENGINE, FALSE TO USE GUI
"""
useStkEngine = False
Read_Scenario = False
############################################################################
# Scenario Setup
############################################################################

if useStkEngine:
    # Launch STK Engine
    print("Launching STK Engine...")
    stkxApp = CreateObject("STKX11.Application")

    # Disable graphics. The NoGraphics property must be set to true before the root object is created.
    stkxApp.NoGraphics = True

    # Create root object
    stkRoot = CreateObject('AgStkObjects11.AgStkObjectRoot')

else:
    # Launch GUI
    print("Launching STK...")
    if not Read_Scenario:
        uiApp = CreateObject("STK11.Application")
    else:
        uiApp = GetActiveObject("STK11.Application")
    uiApp.Visible = True
    uiApp.UserControl = True
    # Get root object
    stkRoot = uiApp.Personality2

# Set date format
stkRoot.UnitPreferences.SetCurrentUnit("DateFormat", "UTCG")
# Create new scenario
print("Creating scenario...")

if not Read_Scenario:
    stkRoot.NewScenario('StarLink')
    
scenario = stkRoot.CurrentScenario
scenario2 = scenario.QueryInterface(STKObjects.IAgScenario)

totalTime = time.time() - startTime
splitTime = time.time()
print("--- Scenario creation: {a:4.3f} sec\t\tTotal time: {b:4.3f} sec ---".format(a=totalTime, b=totalTime))


# 设置开始和结束时间
scenario2.StartTime = '14 Aug 2024 00:00:00.00'
scenario2.StopTime = '14 Aug 2024 00:10:00.00'

# 更新场景时间设置后，应用此时间段
scenario2.Animation.StartTime = scenario2.StartTime
scenario2.Animation.EnableAnimCycleTime = True
scenario2.Animation.AnimCycleTime = scenario2.StopTime

# 设置分析时间步长 
analysis_interval = scenario2.AnalysisInterval
analysis_interval.Step = 6  # 时间步长为1秒

# 设置动画时间步长 
animation = scenario2.Animation
animation.AnimStepValue = 6  # 动画时间步长为1秒
animation.AnimCycleType = 1  # 使用数值1代替 eAnimationCycleStartStop
animation.RefreshDelta = 6  # 每0.5秒刷新一次动画

start_time = scenario2.StartTime
stop_time = scenario2.StopTime

print(f"Scenario Start Time: {start_time}")
print(f"Scenario Stop Time: {stop_time}")

output_file = 'elevation_range.csv'

###########################################################################
# Functions
###########################################################################

           
def CreateUsers(UE_num, name='User'):
    for UE in range(UE_num):
        # 创建目标对象
        target = scenario.Children.New(comtypes.gen.STKObjects.eTarget, f"{name}_{UE}")
        target2 = target.QueryInterface(comtypes.gen.STKObjects.IAgTarget)

        # 随机生成经纬度
        Latitude = random.uniform(39.26, 41.03)  # 纬度
        Longitude = random.uniform(115.25, 117.30)  # 经度

        # 分配目标的地理位置（纬度、经度、高度）
        target2.Position.AssignGeodetic(Latitude, Longitude, 0)

        # 为目标添加传感器
        sensor = target.Children.New(comtypes.gen.STKObjects.eSensor, 'Sensor_' + f"{name}_{UE}")
        sensor2 = sensor.QueryInterface(comtypes.gen.STKObjects.IAgSensor)
        sensor2.SetPatternType(comtypes.gen.STKObjects.eSnSimpleConic)
        sensor2.CommonTasks.SetPatternSimpleConic(90.0, 5.0)

        # 为目标添加接收机
        receiver = target.Children.New(comtypes.gen.STKObjects.eReceiver, 'Receiver_' + f"{name}_{UE}")         
        receiver2 = receiver.QueryInterface(STKObjects.IAgReceiver)
        receiver2.SetModel('Complex Receiver Model')
        recModel = receiver2.Model
        recModel.AutoTrackFrequency = False
        recModel.Frequency = 2
        #gain = recModel.PreReceiveGainsLosses.Add(-2) # dB
        #gain.Identifier = 'Tx and Rx losses'
    
# 创建卫星星系
def Creat_satellite(numOrbitPlanes=22, numSatsPerPlane=72, hight=550, Inclination=53, name='Sat'):
    
    ## numOrbitPlanes 轨道平面数
    ## numSatsPerPlane 每个轨道平面的卫星数量
    ## hight 卫星的轨道高度
    ## Inclination 轨道倾角
    ## name 卫星的名称前缀
    
    # 创建星座对象
    
    constellation = scenario.Children.New(STKObjects.eConstellation, name)
    constellation2 = constellation.QueryInterface(STKObjects.IAgConstellation)

    # 开始插入卫星星座 
    for orbitPlaneNum in range(numOrbitPlanes):  # 使用for循环遍历所有轨道平面，orbitPlaneNum 代表当前的轨道平面索引。这个循环的范围是 numOrbitPlanes，即轨道平面数

        for satNum in range(numSatsPerPlane):  # 嵌套的for循环用于遍历每个轨道平面中的卫星。satNum 表示当前卫星的索引，循环的范围是 numSatsPerPlane，即每个轨道平面的卫星数量
            # 插入卫星
            satellite = scenario.Children.New(STKObjects.eSatellite, f"{name}{orbitPlaneNum}_{satNum}") # 在当前场景中创建一个新的卫星对象，名称由name加上轨道平面和卫星索引组成，例如 Sat0_0
            satellite2 = satellite.QueryInterface(STKObjects.IAgSatellite) # 通过 QueryInterface 方法获取卫星对象的接口 IAgSatellite，用于后续操作

            # 选择传播器
            satellite2.SetPropagatorType(STKObjects.ePropagatorTwoBody) # 设置卫星的传播器类型为 ePropagatorTwoBody，这是 STK 中的一种传播器类型，使用两体动力学来计算卫星轨道
            twoBodyPropagator = satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody) # 获取卫星的两体传播器接口 IAgVePropagatorTwoBody
            
            # 设置初始状态
            keplarian = twoBodyPropagator.InitialState.Representation.ConvertTo(
                STKUtil.eOrbitStateClassical).QueryInterface(STKObjects.IAgOrbitStateClassical)  # 将初始状态表示转换为经典轨道状态（Keplerian elements），使用经典的轨道六要素来描述卫星的轨道
            
            # 设置轨道大小和形状
            keplarian.SizeShapeType = STKObjects.eSizeShapeSemimajorAxis # 设置轨道大小形状类型为半长轴（Semi-Major Axis）
            keplarian.SizeShape.QueryInterface(
                STKObjects.IAgClassicalSizeShapeSemimajorAxis).SemiMajorAxis = hight + 6371  #半长轴设置为轨道高度 hight 加上地球半径6371公里
            keplarian.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeSemimajorAxis).Eccentricity = 0 # 偏心率设置为 0，表示圆形轨道
            
            # 设置轨道倾角和其他轨道要素
            keplarian.Orientation.Inclination = int(Inclination)  # 设置轨道的倾角
            keplarian.Orientation.ArgOfPerigee = 0  # 设置近地点幅角为0度
            keplarian.Orientation.AscNodeType = STKObjects.eAscNodeRAAN # 设置升交点类型为RAAN（升交点赤经）
            
            # 计算和设置RAAN（升交点赤经）
            RAAN = 360 / numOrbitPlanes * orbitPlaneNum # 根据轨道平面数和当前轨道平面的索引计算 RAAN 的值
            keplarian.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = RAAN  # 将计算出的 RAAN 值赋给 AscNode
            
            # 设置位置类型和真近点角
            keplarian.LocationType = STKObjects.eLocationTrueAnomaly # 设置位置类型为真近点角
            trueAnomaly = 360 / numSatsPerPlane * satNum    # 计算每个卫星在其轨道平面内的真近点角
            keplarian.Location.QueryInterface(STKObjects.IAgClassicalLocationTrueAnomaly).Value = trueAnomaly # 将计算的真近点角赋值给 Location

            # 传播轨道
            satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody).InitialState.Representation.Assign(
                keplarian) # 将设置好的经典轨道状态赋给卫星的传播器
            satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody).Propagate() # 调用 Propagate() 方法来传播卫星的轨道，计算出轨道数据

            # 将卫星添加到星座
            constellation2.Objects.AddObject(satellite)
            
# 检查卫星是否与用户有接入；若没有，则直接删除该卫星
def check_and_unload_satellites(scenario):
    # 获取所有的目标（用户）和卫星对象
    target_list = scenario.Children.GetElements(comtypes.gen.STKObjects.eTarget)
    sat_list = scenario.Children.GetElements(comtypes.gen.STKObjects.eSatellite)

    # 遍历每颗卫星和每个用户
    for sat in sat_list:
        sat_name = sat.InstanceName
        sat_accessed = False

        for target in target_list:
            target_name = target.InstanceName

            # 创建Access对象，计算卫星与目标之间的访问情况
            access = sat.GetAccessToObject(target)
            access.ComputeAccess()

            # 检查是否有访问，如果有，标记为True
            if access.ComputedAccessIntervalTimes.Count > 0:
                sat_accessed = True

        # 如果该卫星没有被任何一个用户访问，则将其卸载
        if not sat_accessed:
            scenario.Children.Unload(comtypes.gen.STKObjects.eSatellite, sat_name)
            #print(f"Satellite {sat_name} has been unloaded due to no access.")

    print("Access check done!")

# 拆除用户的传感器（不便于观察 3D 图）
def unload_sensor_from_users(scenario):
    # 获取所有的目标（用户）和卫星对象
    target_list = scenario.Children.GetElements(comtypes.gen.STKObjects.eTarget)
    sat_list = scenario.Children.GetElements(comtypes.gen.STKObjects.eSatellite)

    # 拆除用户的传感器
    for target in target_list:
        sensors = target.Children.GetElements(comtypes.gen.STKObjects.eSensor)
        for sensor in sensors:
            sensor_name = sensor.InstanceName
            target.Children.Unload(comtypes.gen.STKObjects.eSensor, sensor_name)

# 为每个卫星加上发射机和天线
def Add_transmitter_receiver(frequency=2, EIRP=58): # 接受参数 sat_list，这个参数应该是一个卫星对象的列表，每个卫星对象代表STK场景中的一个卫星
    sat_list = scenario.Children.GetElements(comtypes.gen.STKObjects.eSatellite)    
    for each in sat_list: # 遍历卫星列表
        Instance_name = each.InstanceName # 获取当前卫星对象的实例名称，并将其存储在变量 Instance_name 中；这通常是卫星的名称，如在上一个函数中定义的名字
        
        # 创建发射机
        transmitter = each.Children.New(STKObjects.eTransmitter, "Transmitter_" + Instance_name)
        transmitter2 = transmitter.QueryInterface(STKObjects.IAgTransmitter)  # 获取发射机接口，通过 QueryInterface 方法将发射机对象转换为 IAgTransmitter 接口，这样可以访问和设置发射机的具体属性
        transmitter2.SetModel('Simple Transmitter Model') # 设置发射机的模型为 Simple Transmitter Model，这是 STK 中的一个简单模型，适合用于基础的发射机参数设置
        txModel = transmitter2.Model # 获取发射机的模型对象
        txModel = txModel.QueryInterface(STKObjects.IAgTransmitterModelSimple) # 将模型对象转换为 IAgTransmitterModelSimple 接口，这个接口允许设置发射机的简单模型参数
        txModel.Frequency = frequency  # 将发射机的频率设置为传入的参数 frequency
        txModel.EIRP = EIRP  # 设置发射机的等效全向辐射功率（EIRP），单位是 dBW
        #gain = txModel.PostTransmitGainsLosses.Add(-2)
        #txModel.DataRate = DataRate  # 设置发射机的数据传输速率，单位是Mb/sec
        
        # 创建天线
        antenna = each.Children.New(STKObjects.eAntenna, "Antenna_" + Instance_name)
        antenna2 = antenna.QueryInterface(STKObjects.IAgAntenna)  
        antenna2.SetModel('Phased Array')  # 设置为相控阵天线模型
        antModel = antenna2.Model.QueryInterface(STKObjects.IAgAntennaModelPhasedArray)
        antModel.Frequency = frequency  # 设置天线的工作频率
        antModel.Gain = 38 # 天线增益
        

def output_access_data(scenario, output_file):
    # 获取所有的目标和卫星对象
    target_list = scenario.Children.GetElements(comtypes.gen.STKObjects.eTarget)
    sat_list = scenario.Children.GetElements(comtypes.gen.STKObjects.eSatellite)

    # 打开CSV文件进行写入
    with open(output_file, mode='w', newline='') as file:
        writer = csv.writer(file)
        # 写入表头
        writer.writerow(["Target", "Satellite", "Start Time (EpSec)", "End Time (EpSec)", 
                         "Elevation Angle (deg)", "Range (km)"])

        # 遍历每个卫星和每个目标
        for sat in sat_list:
            sat_name = sat.InstanceName
            for target in target_list:
                target_name = target.InstanceName

                # 计算访问
                access = sat.GetAccessToObject(target)
                access.ComputeAccess()

                # 获取访问时间段
                accessIntervals = access.ComputedAccessIntervalTimes
                if accessIntervals.Count > 0:
                    for i in range(accessIntervals.Count):
                        interval = accessIntervals.GetInterval(i)
                        start_time = interval[0]  # 获取开始时间
                        end_time = interval[1]
                        
                        # 获取仰角和距离数据
                        dp = access.DataProviders.Item("AER Data")
                        dp2 = dp.QueryInterface(STKObjects.IAgDataProviderGroup)
                        aer_data = dp2.Group.Item("Default").QueryInterface(comtypes.gen.STKObjects.IAgDataPrvTimeVar)
                        
                        Elements = ["Time", "Elevation", "Range"]
                        results = aer_data.ExecElements(start_time, end_time, 1, Elements)

                        # 提取时间、仰角和距离数据
                        times = results.DataSets.GetDataSetByName("Time").GetValues()
                        elevations = results.DataSets.GetDataSetByName("Elevation").GetValues()
                        ranges = results.DataSets.GetDataSetByName("Range").GetValues()

                        # 写入CSV文件
                        for j in range(len(times)):
                            writer.writerow([target_name, sat_name, times[j], end_time, elevations[j], ranges[j]])

                        #print(f"Access data recorded for Target {target_name} and Satellite {sat_name}.")

                else:
                    print(f"No access between {sat_name} and {target_name}.")

    print(f"Access data recorded in {output_file}!")
    
        
###########################################################################
# Simulation
###########################################################################

if not Read_Scenario:
    
    CreateUsers(UE_num= 10, name='UE') # 创建用户
    Creat_satellite(numOrbitPlanes=22, numSatsPerPlane = 72, hight=550, Inclination=53)  # 创建卫星星座
        
    check_and_unload_satellites(scenario)
    unload_sensor_from_users(scenario)
    Add_transmitter_receiver(frequency=3.5, EIRP=36.7)
    output_access_data(scenario, output_file)

    


