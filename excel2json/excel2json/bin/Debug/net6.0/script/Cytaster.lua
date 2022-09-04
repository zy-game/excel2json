---@class Cytaster 星球表
local Cytaster = {}
local this = Cytaster
---@type number 表示星球的ID
this.CytasterID = nil
---@type string 星球的模型名称
this.CytasterMole = nil
---@type table 星球的贴图名称，可能有多个，配置多少个表示有多少个
this.CytasterMap = nil
---@type number 表示是否为共用星球（0表示：不是共用星球；1表示：是共用星球）
this.CytasterIsPublic = nil
---@type number 购买星球时的权重配比
this.CytasterBuyWeight = nil
---@type number 购买盲盒时的权重配比
this.CytasterBlindWeight = nil
---@type string 星球的初始装饰物占用数据文件配置
this.CytasterInitialData = nil
