--[[
**********************************************************************
LibMagicUtil-1.0 is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

LibMagicUtil is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
General Public License for more details.

You should have received a copy of the GNU General Public License
along with LibMagicUtil.  If not, see <http://www.gnu.org/licenses/>.
**********************************************************************

To use this addon, embed it in your addon using the normal Ace-method
or manually like this:

  LibStub("LibMagicUtil-1.0"):Embed(MyAddon)

After embedding you will be able to use the following methods:

]]

local MAJOR = "LibMagicUtil-1.0"
local MINOR = tonumber("@project-date-integer@") or tonumber(date("%Y%m%d%H%M%S"))
-- Standalone addon sets this flag via bootstrap.lua (loaded from .toc before lib.xml).
-- Embedded copies don't run bootstrap.lua, so the flag is nil for them.
-- Bumping MINOR ensures the standalone copy always wins over embedded copies.
if LIBMAGICUTIL_STANDALONE then
    MINOR = MINOR + 1
end
local lib = LibStub:NewLibrary(MAJOR, MINOR)
local media = LibStub("LibSharedMedia-3.0")
local L = LibStub("AceLocale-3.0"):GetLocale("LibMagicUtil-1.0", false)

if not lib then return end

local embeddables = {
   "GetConfigTemplate", "_GetColorOpt", "_SetColorOpt", "_SetOption", "_GetOption",
   "_HideOption", "_DisableTiling", "FixBackdrop", "_SetBackgroundOption",
   "InterfaceOptions_AddCategory", "InterfaceOptionsFrame_OpenToCategory", "Dump",
   "GetSpellInfo", "UnitAura", "GetMailFrame"
}

lib.optionTemplates = lib.optionTemplates or {}
lib.embeddedAddons = lib.embeddedAddons or {}

local optionTemplates = lib.optionTemplates
local embeddedAddons = lib.embeddedAddons

lib.GetSpellInfo = GetSpellInfo or function(id)
    local spellInfo = C_Spell.GetSpellInfo(id)
    if spellInfo then
        -- Returning the values in the same order as the original GetSpellInfo
        return spellInfo.name, nil, spellInfo.iconID, spellInfo.castTime, spellInfo.minRange, spellInfo.maxRange, spellInfo.spellID, spellInfo.originalIconID
    else
        -- If the spell is not found, return nil
        return nil
    end
end

lib.UnitAura = UnitAura or function(unit, index, filter)
    local aura = C_UnitAuras and C_UnitAuras.GetAuraDataByIndex(unit, index)
    if not aura then return nil end

    -- Map new API fields to old UnitAura return values
    local name = aura.name
    local icon = aura.icon
    local count = aura.applications or 1
    local debuffType = aura.dispelName
    local duration = aura.duration
    local expirationTime = aura.expirationTime
    local caster = aura.sourceUnit
    local isStealable = aura.isStealable
    local nameplateShowPersonal = aura.nameplateShowPersonal
    local spellId = aura.spellId
    local canApplyAura = aura.canApplyAura
    local isBossDebuff = aura.isBossDebuff
    local isCastByPlayer = aura.isFromPlayerOrPlayerPet
    local nameplateShowAll = aura.nameplateShowAll
    local timeMod = aura.timeMod or 1
    local value1, value2, value3 = aura.points or 0, nil, nil

    return name, icon, count, debuffType, duration, expirationTime, caster,
           isStealable, nameplateShowPersonal, spellId, canApplyAura, isBossDebuff,
           isCastByPlayer, nameplateShowAll, timeMod, value1, value2, value3
  end

------------------------------------------------------------------------
-- Reputation API compat
-- Provides modern C_Reputation-style table APIs on all WoW versions.
-- On retail (11.0+), these are thin pass-throughs.
-- On classic/TBC/MoP, they wrap old tuple-returning globals.
------------------------------------------------------------------------

lib.Reputation = {}
local Rep = lib.Reputation

if C_Reputation and C_Reputation.GetFactionDataByIndex then
    -- Retail 11.0+: modern APIs exist natively
    Rep.GetFactionDataByIndex = C_Reputation.GetFactionDataByIndex
    Rep.GetFactionDataByID = C_Reputation.GetFactionDataByID
    Rep.GetNumFactions = C_Reputation.GetNumFactions
    Rep.ExpandFactionHeader = C_Reputation.ExpandFactionHeader
    Rep.CollapseFactionHeader = C_Reputation.CollapseFactionHeader
    Rep.IsFactionActive = C_Reputation.IsFactionActive
    Rep.SetFactionActive = C_Reputation.SetFactionActive
    -- 12.0: IsFactionParagon changed meaning. IsFactionParagonForCurrentPlayer
    -- is "player is at paragon", IsFactionParagon is "faction supports paragon".
    Rep.IsFactionParagon = C_Reputation.IsFactionParagonForCurrentPlayer or C_Reputation.IsFactionParagon
    Rep.GetFactionParagonInfo = C_Reputation.GetFactionParagonInfo
    Rep.IsMajorFaction = C_Reputation.IsMajorFaction
    Rep.IsAccountWideReputation = C_Reputation.IsAccountWideReputation
else
    -- Classic/TBC/MoP: wrap old globals into modern table format
    Rep.GetNumFactions = GetNumFactions or function() return 0 end

    Rep.GetFactionDataByIndex = function(factionIndex)
        if not factionIndex or not GetFactionInfo then return nil end
        local name, description, standingId, barMin, barMax, barValue,
              atWarWith, canToggleAtWar, isHeader, isCollapsed,
              hasRep, isWatched, isChild, factionID,
              hasBonusRepGain, canSetInactive = GetFactionInfo(factionIndex)
        if not name then return nil end
        return {
            factionID = factionID,
            name = name,
            description = description,
            reaction = standingId,
            currentReactionThreshold = barMin,
            nextReactionThreshold = barMax,
            currentStanding = barValue,
            atWarWith = atWarWith,
            canToggleAtWar = canToggleAtWar,
            isHeader = isHeader,
            isCollapsed = isCollapsed,
            isHeaderWithRep = hasRep,
            isWatched = isWatched,
            isChild = isChild,
            hasBonusRepGain = hasBonusRepGain,
            canSetInactive = canSetInactive,
            isAccountWide = false,
        }
    end

    -- Classic doesn't have GetFactionDataByID; return nil
    Rep.GetFactionDataByID = function(factionID)
        return nil
    end

    Rep.ExpandFactionHeader = ExpandFactionHeader or function() end
    Rep.CollapseFactionHeader = CollapseFactionHeader or function() end

    Rep.IsFactionActive = function(factionIndex)
        if IsFactionInactive then
            return not IsFactionInactive(factionIndex)
        end
        return true
    end

    Rep.SetFactionActive = function(factionIndex, setActive)
        if setActive then
            if SetFactionActive then SetFactionActive(factionIndex) end
        else
            if SetFactionInactive then SetFactionInactive(factionIndex) end
        end
    end

    -- No paragon/major faction on classic
    Rep.IsFactionParagon = function() return false end
    Rep.GetFactionParagonInfo = function() return nil end
    Rep.IsMajorFaction = function() return false end
    Rep.IsAccountWideReputation = function() return false end
end

-- Friendship reputation: C_GossipInfo on retail, global on classic
if C_GossipInfo and C_GossipInfo.GetFriendshipReputation then
    Rep.GetFriendshipReputation = function(factionID)
        if not factionID then return nil end
        local rep = C_GossipInfo.GetFriendshipReputation(factionID)
        if not rep or rep.friendshipFactionID == 0 then return nil end
        return rep
    end
else
    Rep.GetFriendshipReputation = function(factionID)
        if not factionID or not GetFriendshipReputation then return nil end
        local friendshipFactionID, standing, maxRep, name, text,
              texture, reaction, reactionThreshold, nextThreshold,
              reversedColor, overrideColor = GetFriendshipReputation(factionID)
        if not friendshipFactionID or friendshipFactionID == 0 then return nil end
        return {
            friendshipFactionID = friendshipFactionID,
            standing = standing,
            maxRep = maxRep,
            name = name,
            text = text,
            texture = texture,
            reaction = reaction,
            reactionThreshold = reactionThreshold,
            nextThreshold = nextThreshold,
            reversedColor = reversedColor,
            overrideColor = overrideColor,
        }
    end
end

if C_GossipInfo and C_GossipInfo.GetFriendshipReputationRanks then
    Rep.GetFriendshipReputationRanks = function(factionID)
        if not factionID then return nil end
        return C_GossipInfo.GetFriendshipReputationRanks(factionID)
    end
else
    Rep.GetFriendshipReputationRanks = function(factionID)
        if not factionID or not GetFriendshipReputationRanks then return nil end
        local currentLevel, maxLevel = GetFriendshipReputationRanks(factionID)
        if not currentLevel then return nil end
        return {
            currentLevel = currentLevel,
            maxLevel = maxLevel,
        }
    end
end

function lib:GetConfigTemplate(config, get, set)
   assert(self ~= lib, "GetConfigTemplate can only be called when embedded.")
   assert(optionTemplates[config] ~= nil, "Unknown config template: "..config)

   local template = {}
   for key, val in pairs(optionTemplates[config]) do
      if key == "args" then
	 local a = {}
	 for akey, aval in pairs(val) do
	    a[akey] = aval
	 end
	 template.args = a
      else
	 template[key] = val
      end
   end
   template.handler = self
   if get then template.get = get end
   if set then template.set = set end
   return template
end

function lib:Embed(addon)
   for _,key in ipairs(embeddables) do
      addon[key] = lib[key]
   end
   embeddedAddons[addon] = true
end

local function GetConfigHash(mod)
   local db = mod.db
   if db then
      if db.profile then
	 return db.profile
      else
	 return db
      end
   elseif mod.GetConfigTable then
      return mod:GetConfigTable()
   end
   assert(db ~= nil, "Couldn't find config hash.")
end

function lib:_SetColorOpt(arg, r, g, b, a)
   local color = arg[#arg]
   local data = GetConfigHash(self).backdropColors
   assert(data ~= nil, "Didn't find the required config data: self.db.profile.backdropColors")
   data[color][1] = r
   data[color][2] = g
   data[color][3] = b
   data[color][4] = a
   if self.OnOptionChanged then
      self:OnOptionChanged(color, data[color], info)
   end
   self:FixBackdrop()
end

function lib:_GetColorOpt(arg)
   local color = arg[#arg]
   return unpack(GetConfigHash(self).backdropColors[color])
end

-- If the current database has 'nil' for this option, hide it
-- All options need to have a default even if it's false 
function lib:_HideOption(info)
   return GetConfigHash(self)[info[#info]] == nil
end


function lib:_SetOption(info, val)
   local var = info[#info]
   GetConfigHash(self)[info[#info]] = val
   if self.OnOptionChanged then
      self:OnOptionChanged(var, val, info)
   end
end

function lib:_GetOption(info)
   return GetConfigHash(self)[info[#info]]
end

function lib:_DisableTiling(info)
   return not GetConfigHash(self).tile 
end

function lib:FixBackdrop()   
   local bgFrame = self.frame:GetBackdrop()
   local db = GetConfigHash(self)
   if not bgFrame then
      bgFrame = {
	 insets = {left = 1, right = 1, top = 1, bottom = 1}
      }
   end

   local edge = 0
   local inset = 0
   if db.border ~= "None" then
      edge = db.edgeSize
      inset = db.edgeSize / 4
   end
   bgFrame.edgeSize = edge
   bgFrame.insets.left   = inset
   bgFrame.insets.right  = inset
   bgFrame.insets.top    = inset
   bgFrame.insets.bottom = inset

   bgFrame.tile = db.tile
   bgFrame.tileSize = db.tileSize

   bgFrame.edgeFile = media:Fetch("border", db.border)
   bgFrame.bgFile = media:Fetch("background", db.background)

   self.frame:SetBackdrop(bgFrame)
   self.frame:SetBackdropColor(unpack(db.backdropColors.backgroundColor))
   self.frame:SetBackdropBorderColor(unpack(db.backdropColors.borderColor))
end

function lib:_SetBackgroundOption(...)
   self:_SetOption(...)
   self:FixBackdrop()
end

function lib:Dump(o)
    if type(o) == 'table' then
        local s = '{ '
        for k,v in pairs(o) do
            if type(k) ~= 'number' then k = '"'..k..'"' end
            s = s .. '['..k..'] = ' .. lib:Dump(v) .. ','
        end
        return s .. '} '
    else
        return tostring(o)
    end
end




function lib:InterfaceOptions_AddCategory(frame, addOn, position)
    if InterfaceOptions_AddCategory then
        return InterfaceOptions_AddCategory(frame, addon, position)
    end
	-- cancel is no longer a default option. May add menu extension for this.
	frame.OnCommit = frame.okay;
	frame.OnDefault = frame.default;
	frame.OnRefresh = frame.refresh;

	if frame.parent then
		local category = Settings.GetCategory(frame.parent)
		local subcategory = Settings.RegisterCanvasLayoutSubcategory(category, frame, frame.name, frame.name)
		subcategory.ID = frame.name;
		return subcategory, category;
	else
		local category = Settings.RegisterCanvasLayoutCategory(frame, frame.name, frame.name);
		category.ID = frame.name;
		Settings.RegisterAddOnCategory(category);
		return category;
	end
end

-- Deprecated. Use Settings.OpenToCategory().
function lib:InterfaceOptionsFrame_OpenToCategory(categoryIDOrFrame)
    if InCombatLockdown() then
        return
    end
    if InterfaceOptionsFrame_OpenToCategory then
        return InterfaceOptionsFrame_OpenToCategory(categoryIDOrFrame)
    end
	if type(categoryIDOrFrame) == "table" then
		local categoryID = categoryIDOrFrame.name;
		return Settings.OpenToCategory(categoryID);
	else
		return Settings.OpenToCategory(categoryIDOrFrame);
	end
end

-- Returns the active mail frame and whether it's TSM.
-- Second return value is true if the returned frame is TSM's mail frame.
-- Finds TSM's mail window by looking for the MailsScrollTable element
-- (unique to TSM's mail UI) and walking up to the top-level parent.
local cachedTSMMailFrame
local tsmScanComplete = false
local function FindTSMMailFrame()
    if cachedTSMMailFrame then
        return cachedTSMMailFrame
    end
    -- Already scanned this mailbox session and didn't find it
    if tsmScanComplete then
        return nil
    end
    -- Find MailsScrollTable, then walk up to the LargeApplicationFrame
    local frame = EnumerateFrames()
    while frame do
        -- Forbidden frames (e.g. Blizzard aura buttons) error on any method
        -- call from addon code except IsForbidden.
        local name = not frame:IsForbidden() and frame:GetName()
        if name and type(name) == "string" and name:match("^TSM_FRAME:MailsScrollTable:") then
            -- Walk up the parent chain to find the top-level TSM frame
            local parent = frame:GetParent()
            while parent and parent ~= UIParent and not parent:IsForbidden() do
                local parentName = parent:GetName()
                if parentName and type(parentName) == "string" and parentName:match("^TSM_FRAME:LargeApplicationFrame:") then
                    cachedTSMMailFrame = parent
                    tsmScanComplete = true
                    return parent
                end
                parent = parent:GetParent()
            end
        end
        frame = EnumerateFrames(frame)
    end
    tsmScanComplete = true
    return nil
end

-- Reset the negative scan cache so the next GetMailFrame call re-scans.
-- Call this when the mailbox opens, giving TSM a chance to create its frame.
function lib:ResetTSMFrameCache()
    if not cachedTSMMailFrame then
        tsmScanComplete = false
    end
end

function lib:GetMailFrame()
    local tsmFrame = FindTSMMailFrame()
    if tsmFrame and tsmFrame:IsShown() then
        return tsmFrame, true
    end
    return MailFrame, false
end

-- Config template for a frame background
optionTemplates.background = {
   type = "group",
   name = L["Background Frame"],
   set = "_SetOption",
   get = "_GetOption",
   args = {
      background = {
	 type = "select",
	 dialogControl = "LSM30_Background",
	 name = L["Background Texture"],
	 desc = L["The background texture used for the background frame."], 
	 order = 20,
	 values = AceGUIWidgetLSMlists.background,
	 set = "_SetBackgroundOption", 
      },
      backgroundColor = {
	 type = "color",
	 name = L["Background Color"],
	 hasAlpha = true,
	 set = "_SetColorOpt",
	 get = "_GetColorOpt",
	 order = 30,
      },
      spacer1 = {
	 type = "description",
	 width = "full",
	 name = "",
	 order = 35,
	 set = "_SetBackgroundOption", 
      },
      border = {
	 type = "select",
	 dialogControl = "LSM30_Border",
	 name = L["Border Texture"],
	 desc = L["The border texture used for the background frame."],
	 order = 40,
	 values = AceGUIWidgetLSMlists.border, 
	 set = "_SetBackgroundOption", 
      },
      borderColor = {
	 type = "color",
	 name = L["Border color"],
	 hasAlpha = true,
	 set = "_SetColorOpt",
	 get = "_GetColorOpt",
	 order = 50,
      },
      spacer2 = {
	 type = "description",
	 width = "full",
	 name = "",
	 order = 55,
      },
      edgeSize = {
	 type = "range",
	 name = L["Edge size"],
	 desc = L["Width of the border."],
	 min = 1, max = 50, step = 0.1,
	 set = "_SetBackgroundOption",
	 order = 60, 
      },
      padding = {
	 type = "range",
	 name = L["Padding"],
	 desc = L["Number of pixels to insert between the background frame edge and the content."],
	 min = 0, max = 50, step = 0.1,
	 order = 70, 
      },
      tile = {
	 type = "toggle",
	 name = L["Tile Background"],
	 desc = L["Whether or not to tile the background texture."],
	 order = 100,
	 set = "_SetBackgroundOption", 
	 order = 90, 
      },
      tileSize = {
	 type = "range",
	 name = L["Tile Size"],
	 desc = L["The size in pixels of the background tiles."],
	 min = 1, max = 200, step = 0.1, 
	 order = 110,
	 disabled = "_DisableTiling",
	 set = "_SetBackgroundOption", 
	 order = 100,
      },
   }
}

optionTemplates.barsize = {
   type = "group",
   name = L["Bar Size"],
   get = "_GetOption",
   set = "_SetOption",
   args = {
      maxbars  = {
	 type = "range",
	 min = 1, max = 100, step = 1,
	 name = L["Maximum number of bars"],
	 width="full",
	 order = 0,
	 hidden = "_HideOption", 
      },
      height = {
	 type = "range",
	 name = L["Bar Height"],
	 width = "full",
	 min = 1, max = 50, step = 1,
      }, 
      width = {
	 type = "range",
	 name = L["Bar Width"],
	 width = "full",
	 min = 1, max = 500, step = 1,
      }, 
      scale = {
	 type = "range",
	 name = L["Bar Scale"],
	 width = "full",
	 min = 0.01, max = 5, step = 0.005,
      }, 
      spacing = {
	 type = "range",
	 name = L["Bar Spacing"],
	 desc = L["The space in pixels between each bar."],
	 width = "full",
	 min = 0, max = 40, step = 0.05,
	 hidden = "_HideOption", 
      }, 
   }
}

-- upgrade
for addon in pairs(embeddedAddons) do
   lib:Embed(addon)
end
