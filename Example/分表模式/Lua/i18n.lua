---@class Cfg_i18n
---@field language string
---@field Items_name_1 string
---@field Items_desc_1 string
---@field Items_name_2 string
---@field Items_desc_2 string
---@field TempConfig_name_1 string
---@field TempConfig_desc_1 string
---@field TempConfig_name_2 string
---@field TempConfig_desc_2 string
---@field TempConfig_name_3 string
---@field TempConfig_desc_3 string
---@field TempConfig_name_4 string
---@field TempConfig_desc_4 string

---@type Cfg_i18n
local cfg_i18n = {
	language = "en"
}

setmetatable(cfg_i18n, {
	__index = function (t, key)
		local languages = require ("Data." .. t.language)
		if languages[key] == nil then
			CS.UnityEngine.Debug.LogError(languages[key],("多语言表的[%s]语种中不存在key[%s]\n%s"):format(tostring(t.language),tostring(key), debug.traceback()))
		end
		return languages[key]
	end
})

return cfg_i18n
