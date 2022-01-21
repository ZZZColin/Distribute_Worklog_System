# Distribute_Worklog_System

分布式工作表适用于缺少SharePoint权限并规模较大的队伍实现Task Info的同步，规避了使用Excel ShareWorkbook功能带来的频繁同步导致Excel假死造成数据丢失的问题。

在使用前，请打开Master Log中的Log Tab，此Tab由三部分组成：

>1. 环境变量设置
---------------
>>>![Capture](https://user-images.githubusercontent.com/49432881/150459895-a0ec079b-dfc3-4fb5-a8f1-66d182744ab2.PNG)
     
>>>FolderPath：用于存储Individual Log所在文件夹路径
     
>>>Date Column：用于存储分表日期所在列
     
>>>Key Column：用于存储键所在列（键用于查重）
     
>>>Column Count：用于存储读取Individual Log数据范围（如10代表读取Individual Log第一列到第十列）
     
>>>Overwrite String Cell：用于存储Individual Log上Overwrite String所在单元格位置（Overwrite String用于标识需要Overwrite的数据，由行号组成，用空格分隔不同行号）
    
>2. Update和Overwrite功能区
----------------  
>>>![Capture1](https://user-images.githubusercontent.com/49432881/150459919-5f69cdf9-dcf7-47ab-a936-8fd58c6cfe28.PNG)
      
>>>Update：由日期范围及运行按钮组成，使用前需输入起止日期
     
>>>Overwrite：点击运行按钮即可使用
      
>3. 运行Log
---------------- 
>>>用于显示运行过程中产生的中间步骤，对于Update和Overwrite，Log可分为：
     
>>>>Update：
     
>>>>>* Update起始行：用于标识Update开始时Master Log的起始行号，如Update Data From Row 10
        
>>>>>* 文件路径：当前读取的文件完全路径
        
>>>>>* 文件修改日期：当前读取的文件的最新修改日期（用于判断程序读取文件版本是否最新。当Individual Log位于公共磁盘时，可能出现文件保存不同步的问题）
        
>>>>>* 读取数据范围：用于标识读取文件的数据范围，由环境变量中的Column Count提供列数，Individual Log中日期所在列中从下往上第一个非空单元格所在行决定行数，如Data Range: A1:J100
        
>>>>>* 逐条信息：用于显示每条数据的Update情况

>>>>>>>成功，如Row: 10, Value: 123 Updated On Row 15

>>>>>>>键重复，如Row: 10, Value: 123 Duplicate On Row 15

>>>>>* 汇总信息：用于显示当前Individual Log Update的总数，如10 Record(s) Updated In Total.
        
        
