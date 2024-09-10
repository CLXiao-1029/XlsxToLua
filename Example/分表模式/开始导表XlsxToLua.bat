::程序名
::设置 Excel 目录					../../../../../Config
::设置 导出文件 目录				../../../Game/Scripts/LuaScripts/Shared/Data
::设置 导出文件 是否有序			false
::是否生成总表						true		总表模式需要在Excel目录下新建子文件夹，并将其配置放到子文件夹中
::是否拆分导出文件的注解			true
::文本提取							true		摘取表格中的文本并转换成翻译的key。默认true开启									
::显示时间流逝						true
::获取提交记录						1	0:默认不获取 1:Git 2:Svn
::导出文件名字规则					0 总表模式有效 0:默认 1:首字母大写 2:首字母小写 3:转大写 4:转小写 5:大驼峰命名 6:小驼峰命名

".\XlsxToLua.exe" ./Config ./Lua false false true true true 0 0



::"E:\Working\Unity\Develop\RoamingFortress\RoamingFortress\Assets\Engine\Pack~\Config"
::"E:\Working\Unity\Develop\RoamingFortress\Config"
::"E:\Working\Unity\Develop\RoamingFortress\RoamingFortress\Assets\Game\Scripts\LuaScripts\Shared\Data"
pause