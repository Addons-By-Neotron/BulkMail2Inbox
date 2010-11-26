BulkMailInbox = LibStub("AceAddon-3.0"):NewAddon("BulkMailInbox", "AceConsole-3.0", "AceEvent-3.0", "AceTimer-3.0", "AceHook-3.0")

local mod, self, BulkMailInbox = BulkMailInbox, BulkMailInbox, BulkMailInbox

local VERSION = "4.0-beta"
local LibStub = LibStub

local L        = LibStub("AceLocale-3.0"):GetLocale("BulkMailInbox", false)

BulkMailInbox.L = L

local abacus   = LibStub("LibAbacus-3.0")
local QTIP     = LibStub("LibQTip-1.0")
local AC       = LibStub("AceConfig-3.0")
local ACD      = LibStub("AceConfigDialog-3.0")
local DB       = LibStub("AceDB-3.0")
local LDB      = LibStub("LibDataBroker-1.1", true)
local LD       = LibStub("LibDropdown-1.0")


local _G = _G
local fmt = string.format
local lower = string.lower

local sortFields, markTable  -- tables
local ibIndex, ibAttachIndex, numInboxItems, inboxCash, cleanPass, cashOnly, markOnly, takeAllInProgress, invFull, filterText -- variables

--[[----------------------------------------------------------------------------
Table Handling
------------------------------------------------------------------------------]]
local newHash, del
do
   local list = setmetatable({}, {__mode='k'})
   function newHash(...)
      local t = next(list)
      if t then
	 list[t] = nil
      else
	 t = {}
      end
      for i = 1, select('#', ...), 2 do
	 t[select(i, ...)] = select(i+1, ...)
      end
      return t
   end
   
   function del(t)
      for k in pairs(t) do
	 t[k] = nil
      end
      list[t] = true
      return nil
   end
end

--[[----------------------------------------------------------------------------
Local Processing
------------------------------------------------------------------------------]]
-- Build a table with info about all items and money in the Inbox
local inboxCache = {}
local function _sortInboxFunc(a,b)
   local sf = sortFields[BulkMailInbox.db.char.sortField]
   if a and b then
      a, b = a[sf], b[sf]
      if sf == 'itemLink' then
	 a = GetItemInfo(a) or a
	 b = GetItemInfo(b) or b
      end
      a = type(a) == "nil" and 0 or type(a) == "boolean" and tostring(a) or a
      b = type(b) == "nil" and 0 or type(b) == "boolean" and tostring(b) or b
      if mod.db.char.sortReversed then
	 if a > b then return true end
      else
	 if a < b then return true end
      end
   end
end

-- These are the patterns that indicate that the item was received from the AH
-- In that case the items are not returnable
local AHReceivedPatterns = {
   (gsub(AUCTION_REMOVED_MAIL_SUBJECT, "%%s", ".*")), 
   (gsub(AUCTION_EXPIRED_MAIL_SUBJECT, "%%s", ".*")), 
   (gsub(AUCTION_WON_MAIL_SUBJECT, "%%s", ".*"))
}

local function _isAHSentMail(subject)
   if subject then
      for key,pattern in ipairs(AHReceivedPatterns) do
	 if subject:find(pattern) then
	    return true
	 end
      end
   end
end

local function _matchesFilter(text)
   if not filterText or filterText:len() == 0 then
      return true
   end
   return text:lower():find(filterText, 1, true)
end

local function inboxCacheBuild()
   local start = GetTime()
   for k in ipairs(inboxCache) do inboxCache[k] = del(inboxCache[k]) end
   inboxCash, numInboxItems = 0, 0
   for i = 1, GetInboxNumItems() do
      local _, _, sender, subject, money, cod, daysLeft, numItems, _, wasReturned, _, canReply, isGM = GetInboxHeaderInfo(i)
      if money > 0 then
	 local title = itemName and ITEM_SOLD_COLON..' '..itemName or L["Cash"]
	 if _matchesFilter(title) then
	    -- Contributed by Scott Centoni
	    local _, itemName = GetInboxInvoiceInfo(i)
	    table.insert(inboxCache, newHash(
			    'index', i, 'sender', sender, 'bmid', daysLeft..subject..0, 'returnable', not wasReturned, 'cod', cod,
			    'daysLeft', daysLeft, 'itemLink', itemName and ITEM_SOLD_COLON..' '..itemName or L["Cash"], 'money', money, 'texture', "Interface\\Icons\\INV_Misc_Coin_01"
		      ))
	    inboxCash = inboxCash + money
	 end
      end
      if numItems then
	 local canReturnItem = not wasReturned
	 if isGM or not canReply or _isAHSentMail(subject) then
	    canReturnItem = false
	 end
	 for j=1, ATTACHMENTS_MAX_SEND do
	    if GetInboxItem(i,j) and _matchesFilter(GetInboxItem(i, j)) then
	       table.insert(inboxCache, newHash(
			       'index', i, 'attachment', j, 'sender', sender, 'bmid', daysLeft..subject..j, 'returnable', canReturnItem, 'cod', cod,
			       'daysLeft', daysLeft, 'itemLink', GetInboxItemLink(i,j), 'qty', select(3, GetInboxItem(i,j)), 'texture', (select(2, GetInboxItem(i,j)))
			 ))
	       numInboxItems = numInboxItems + 1
	    end
	 end
      end
   end
   table.sort(inboxCache, _sortInboxFunc)
end

local function takeAll(cash, mark)
   -- Ace3 timers only allow one arg so support that hack
   if type(cash) == "table" then
      mark = cash.markOnly
      cash = cash.cashOnly
   end
   
   cashOnly = cash
   markOnly = mark
   ibIndex = GetInboxNumItems()
   ibAttachIndex = 0
   takeAllInProgress = true
   inboxCacheBuild()
   mod:TakeNextItemFromMailbox()
end

--[[----------------------------------------------------------------------------
Setup
------------------------------------------------------------------------------]]
local function color(text, color)
   return fmt("|cff%s%s|r", color, text)
end

function mod:OnInitialize()
   if not BulkMail3InboxDB and BulkMail2InboxDB and BulkMail2InboxDB.chars then
      BulkMail3InboxDB = { char = {} }
      for charname, data in pairs(BulkMail2InboxDB) do
	 BulkMail3InboxDB.char[charname] = data
      end
   end
   self.db = DB:New("BulkMail3InboxDB", {
		       char = {
			  altDel = false,
			  ctrlRet = true,
			  shiftTake = true,
			  takeAll = true,
			  inboxUI = true,
			  sortField = 1,
		       },
		    }, "Default")
   
   sortFields = { 'itemLink', 'qty', 'returnable', 'sender', 'daysLeft', 'index' }
   markTable = {}
   inboxCash = 0
   invFull = false

   self.opts = {
      type = 'group',
      args = {
	 altdel = {
	    name = L["Alt-click Delete"], type = 'toggle', 
	    desc = L["Enable Alt-Click on inbox items to delete the mail in which they are contained."],
	    get = function() return self.db.char.altDel end,
	    set = function(args,v) self.db.char.altDel = v end,
	 },
	 ctrlret = {
	    name = L["Ctrl-click Return"], type = 'toggle', 
	    desc = L["Enable Ctrl-click on inbox items to return the mail in which they are contained."],
	    get = function() return self.db.char.ctrlRet end,
	    set = function(args,v) self.db.char.ctrlRet = v end,
	 },
	 shifttake = {
	    name = L["Shift-click Take"], type = 'toggle',
	    desc = L["Enable Shift-click on inbox items to take them."],
	    get = function() return self.db.char.shiftTake end,
	    set = function(args,v) self.db.char.shiftTake = v end,
	 },
	 takeall = {
	    name = L["Take All"], type = 'toggle',
	    desc = L["Enable 'Take All' button in inbox."],
	    get = function() return self.db.char.takeAll end,
	    set = function(args,v) self.db.char.takeAll = v; self:UpdateTakeAllButton() end,
	 },
	 gui = {
	    name = L["Show Inbox GUI"], type = 'toggle',
	    desc = L["Show the Inbox Items GUI"],
	    get = function() return self.db.char.inboxUI end,
	    set = function(args,v) self.db.char.inboxUI = v; self:RefreshInboxGUI() end,
	 },
      },
   }

   -- set up LDB, but only if the user doesn't have Bulk Mail already
   if LDB and not BulkMail then
      self.ldb =
	 LDB:NewDataObject("BulkMailInbox",
			   {
			      type =  "launcher", 
			      label = L["Bulk Mail Inbox"]..VERSION,
			      icon = [[Interface\Addons\BulkMail2\icon]],
			      tooltiptext = color(L["Bulk Mail Inbox"]..VERSION.."\n\n", "ffff00")..color(L["Hint:"].." "..L["Left click to open the config panel."].."\n"..
											  L["Right click to open the config menu."], "ffd200"),
			      OnClick = function(clickedframe, button)
					   if button == "RightButton" then
					      mod:OpenConfigMenu(clickedframe)
					   else
					      mod:ToggleConfigDialog()
					   end
					end,
			   })
   end

   self._mainConfig = self:OptReg(L["Bulk Mail Inbox"], self.opts,  { "bmi", "bulkmailinbox" })

   if BulkMail then
      BulkMail.opts.args.inbox = { type = "group",
				   handler = mod,
				   name = L["Inbox"],
				   desc = L["Bulk Mail Inbox Options"],
				   args = BulkMailInbox.opts.args
				}
   end
end

function mod:OnEnable()
   self:UpdateTakeAllButton()
   self:RegisterEvent('MAIL_SHOW')
   self:RegisterEvent('MAIL_CLOSED')
   self:RegisterEvent('PLAYER_ENTERING_WORLD')
   self:RegisterEvent('UI_ERROR_MESSAGE')
   self:RegisterEvent('MAIL_INBOX_UPDATE')

   -- Handle being LoD loaded while at the mailbox
   if MailFrame:IsVisible() then
      self:MAIL_SHOW()
   end
end

function mod:OnDisable()
   self:UnregisterAllEvents()
end

--[[----------------------------------------------------------------------------
Events
------------------------------------------------------------------------------]]
function mod:MAIL_SHOW()
   ibIndex = GetInboxNumItems()

   if not self:IsHooked('CheckInbox') then
      self:SecureHook('CheckInbox', 'RefreshInboxGUI')
      self:SecureHook(GameTooltip, 'SetInboxItem')
      self:Hook('InboxFrame_OnClick', nil, true)
      self:SecureHookScript(MailFrameTab1, 'OnClick', 'ShowInboxGUI')
      self:SecureHookScript(MailFrameTab2, 'OnClick', 'HideInboxGUI')
   end

   self:ShowInboxGUI()
end

function mod:MAIL_CLOSED()
   takeAllInProgress = false
   self:HideInboxGUI()
   GameTooltip:Hide()
   self:UnhookAll()
end

BulkMailInbox.PLAYER_ENTERING_WORLD = BulkMailInbox.MAIL_CLOSED  -- MAIL_CLOSED doesn't get called if, for example, the player accepts a port with the mail window open

function mod:UI_ERROR_MESSAGE(event, msg)  -- prevent infinite loop when inventory is full
   if msg == ERR_INV_FULL then
      invFull = true
   end
end

-- Take next inbox item or money; skip past CoD items and letters.
local prevSubject = ''

function mod:SmartCancelTimer(name)
   mod.timers = mod.timers or {}
   if mod.timers[name] then
      mod:CancelTimer(mod.timers[name], true)
      mod.timers[name] = nil
   end
end

function mod:SmartScheduleTimer(name, override, method, timeout, ...)
   mod.timers = mod.timers or {}
   if mod.timers[name] and override then
      mod:CancelTimer(mod.timers[name], true)
      mod.timers[name] = nil
   end
   if not mod.timers[name] then
      mod.timers[name] = mod:ScheduleTimer(method, timeout, ...)
   end
end

function mod:MAIL_INBOX_UPDATE()
   if not takeAllInProgress and not self.refreshInboxTimer then
      mod:SmartScheduleTimer('BMI_RefreshInboxGUI', false, "RefreshInboxGUI", .5)
   end
end

function mod:TakeNextItemFromMailbox()
   if not takeAllInProgress then return end

   local numMails = GetInboxNumItems()
   cashOnly = cashOnly or invFull
   if ibIndex <= 0 then
      if cleanPass or numMails <= 0 then
	 takeAllInProgress = false
	 invFull = false
	 return self:RefreshInboxGUI()
      else
	 ibIndex = numMails
	 ibAttachIndex = 0
	 cleanPass = true
	 return self:SmartScheduleTimer('BMI_takeAll', true, takeAll, .1, { cashOnly = cashOnly, markOnly = markOnly })
      end
   end
   
   local curIndex, curAttachIndex = ibIndex, ibAttachIndex
   local subject, money, cod, daysLeft, item, _, _, text, _, isGM = select(4, GetInboxHeaderInfo(curIndex))

   if subject then
      prevSubject = subject
   else
      subject = prevSubject
   end

   if curAttachIndex == ATTACHMENTS_MAX_SEND then
      ibIndex = ibIndex - 1
      ibAttachIndex = 0
   else
      ibAttachIndex = ibAttachIndex + 1
   end

   if curAttachIndex > 0 and not GetInboxItem(curIndex, curAttachIndex) or markOnly and not markTable[daysLeft..subject..curAttachIndex] then
      return self:TakeNextItemFromMailbox()
   end

   if not string.find(subject, "Sale Pending") then 
      if curAttachIndex == 0 and money > 0 then
	 cleanPass = false
	 TakeInboxMoney(curIndex)
      elseif not cashOnly and cod == 0 then
	 cleanPass = false
	 if not invFull then
	    TakeInboxItem(curIndex, curAttachIndex)
	 end
      end
   end

   if not cleanPass then
      self:SmartScheduleTimer('BMI_RefreshInboxGUI', false, "RefreshInboxGUI", 1)
      self:SmartScheduleTimer('BMI_TakeNextItem', true, "TakeNextItemFromMailbox", 0.4)
      return
   end
   -- Tail recurse
   return self:TakeNextItemFromMailbox()
end

--[[----------------------------------------------------------------------------
Hooks
------------------------------------------------------------------------------]]
function mod:SetInboxItem(tooltip, index, attachment, ...)
   if takeAllInProgress then return end
   local money, _, _, _, _, wasReturned, _, canReply = select(5, GetInboxHeaderInfo(index))
   if self.db.char.shiftTake then tooltip:AddLine(L["Shift - Take Item"]) end
   if wasReturned then 
      if self.db.char.altDel then
	 tooltip:AddLine(L["Alt - Delete Containing Mail"])
      end
   elseif canReply and self.db.char.ctrlRet then
      tooltip:AddLine(L["Ctrl - Return Containing Mail"])
   end
end

function mod:InboxFrame_OnClick(parentself, index, attachment, ...)
   takeAllInProgress = false
   local _, _, _, _, money, cod, _, hasItem, _, wasReturned, _, canReply = GetInboxHeaderInfo(index)
   if self.db.char.shiftTake and IsShiftKeyDown() then
      if money > 0 then TakeInboxMoney(index)
      elseif cod > 0 then return
      elseif hasItem then TakeInboxItem(index, attachment) end
   elseif self.db.char.ctrlRet and IsControlKeyDown() and not wasReturned and canReply then ReturnInboxItem(index)
   elseif self.db.char.altDel and IsAltKeyDown() and wasReturned then DeleteInboxItem(index)
   elseif parentself and parentself:GetObjectType() == 'CheckButton' then self.hooks.InboxFrame_OnClick(parentself, index, ...) end
   mod:SmartScheduleTimer("BMI_RefreshInboxGUI", true, "RefreshInboxGUI", 0.1)
end

--[[----------------------------------------------------------------------------
Inbox GUI
------------------------------------------------------------------------------]]
-- Update/Create the Take All button
function mod:UpdateTakeAllButton()
   if self.db.char.takeAll then
      if _G.BMI_TakeAllButton then return end
      local bmiTakeAllButton = CreateFrame("Button", "BMI_TakeAllButton", InboxFrame, "UIPanelButtonTemplate")
      bmiTakeAllButton:SetWidth(120)
      bmiTakeAllButton:SetHeight(25)
      bmiTakeAllButton:SetPoint("CENTER", InboxFrame, "TOP", -15, -410)
      bmiTakeAllButton:SetText(L["Take All"])
      bmiTakeAllButton:SetScript("OnClick", function() takeAll() end)
   else
      if _G.BMI_TakeAllButton then _G.BMI_TakeAllButton:Hide() end
      _G.BMI_TakeAllButton = nil
   end
end

-- Inbox Items Tablet
local function highlightSameMailItems(index, ...)
   if self.db.char.altDel and IsAltKeyDown() or self.db.char.ctrlRet and IsControlKeyDown() then
      for i = 1, select('#', ...) do
	 local row = select(i, ...)
	 if row.col6 and row.col6:GetText() == index then
	    row.highlight:Show()
	 end
      end
   end
end

local function unhighlightSameMailItems(index, ...)
   for i = 1, select('#', ...) do
      local row = select(i, ...)
      if row.col6 and row.col6:GetText() == index then
	 row.highlight:Hide()
      end
   end
end

--[[----------------------------------------------------------------------------
QTip GUI
------------------------------------------------------------------------------]]

local function _addIndentedCell(tooltip, text, indentation, func, arg)
   local y, x = tooltip:AddLine()
   tooltip:SetCell(y, x, text, tooltip:GetFont(), "LEFT", 1, nil, indentation)
   if func then
      tooltip:SetLineScript(y, "OnMouseUp", func, arg)
   end
   return y, x
end

local function _addColspanCell(tooltip, text, colspan, func, arg, y)
   y = y or tooltip:AddLine()
   tooltip:SetCell(y, 1, text, tooltip:GetFont(), "LEFT", colspan)
   if func then
      tooltip:SetLineScript(y, "OnMouseUp", func, true)
   else
      tooltip:SetLineScript(y, "OnMouseUp", nil)
   end
   return y
end

function mod:HideInboxGUI()
   if mod._toolbar then
      mod._toolbar:Hide()
      mod._toolbar:SetParent(nil)
   end

   local tooltip = mod.inboxGUI
   if tooltip then
      tooltip:EnableMouse(false)
      tooltip:SetScript("OnDragStart", nil)
      tooltip:SetScript("OnDragStop", nil)
      tooltip:SetMovable(false)
      tooltip:RegisterForDrag()
      tooltip:SetFrameStrata("TOOLTIP")
      tooltip.moved = nil
      QTIP:Release(tooltip)
      mod.inboxGUI = nil
   end
   mod._wantGui = nil
end

function mod:RefreshInboxGUI()
   mod:SmartCancelTimer('BMI_RefreshInboxGUI')
   if not mod.db.char.inboxUI then return end
   inboxCacheBuild()
   if mod.inboxGUI then
      -- Rebuild it since it's visible
      mod:ShowInboxGUI()
   end
end


local function _onLeaveFunc(frame, info)
   if mod.tooltipShowing == frame then
      GameTooltip:Hide()
      mod.tooltipShowing = nil
      frame:SetScript("OnKeyUp", nil)
      frame:SetScript("OnKeyDown", nil)
   end
end

local function _toggleCompareItem()
   if IsShiftKeyDown() then
      GameTooltip_ShowCompareItem()
   else
      -- There appears to be no other way. Sigh.
      if ( GameTooltip.shoppingTooltips ) then
	 for _, frame in pairs(GameTooltip.shoppingTooltips) do
	    frame:Hide();
	 end
      end
      GameTooltip.comparing = false;
   end
end

local function _onEnterFunc(frame, info)  -- contributed by bigzero
   mod.tooltipShowing = frame
   GameTooltip:SetOwner(frame, 'ANCHOR_BOTTOMRIGHT', 0, 0)
   if info.index and info.attachment and GetInboxItem(info.index, info.attachment) then
      GameTooltip:SetInboxItem(info.index, info.attachment)
   end
   if IsShiftKeyDown() then
      GameTooltip_ShowCompareItem()
   end
   if info.money then
      GameTooltip:AddLine(ENCLOSED_MONEY, "", 1, 1, 1)
      SetTooltipMoney(GameTooltip, info.money)
      SetMoneyFrameColor('GameTooltipMoneyFrame', HIGHLIGHT_FONT_COLOR.r, HIGHLIGHT_FONT_COLOR.g, HIGHLIGHT_FONT_COLOR.b)
   end 
   if (info.cod or 0) > 0 then
      GameTooltip:AddLine(COD_AMOUNT, "", 1, 1, 1)
      SetTooltipMoney(GameTooltip, info.cod)
      SetMoneyFrameColor('GameTooltipMoneyFrame', HIGHLIGHT_FONT_COLOR.r, HIGHLIGHT_FONT_COLOR.g, HIGHLIGHT_FONT_COLOR.b)
   end 
   GameTooltip:Show()
   frame:SetScript("OnKeyDown", _toggleCompareItem)
   frame:SetScript("OnKeyUp", _toggleCompareItem)
end

local function _createButton(title, parent, onclick, anchorTo, xoffset)
   local buttons = mod.buttons or {}
   mod.buttons = buttons
   
   local button = CreateFrame("Button", nil, parent, "UIPanelButtonTemplate")
   button:SetText(title)
   button:SetWidth(25)
   button:SetHeight(20)
   button:SetScript("OnClick", onclick)
   buttons[title] = button
   button:SetPoint("RIGHT", anchorTo, "LEFT", xoffset, 0)
   return button
end

local function _createOrAttachSearchBar(tooltip)
   local toolbar = mod._toolbar
   if not toolbar then
      toolbar = CreateFrame("Frame", nil, UIParent)
      toolbar:SetHeight(35)

      local button =  CreateFrame("Button", nil, tooltip, "UIPanelCloseButton")
      button:SetPoint("RIGHT", toolbar, "RIGHT", 0, 0)
      button:SetScript("OnClick", function() mod:HideInboxGUI() end)
      
      button = _createButton("CS", toolbar, function() wipe(markTable) self:RefreshInboxGUI() end, button, -2)
      button = _createButton("TS", toolbar, function() takeAll(false, true) end, button, -2)
      button = _createButton("TC", toolbar, function() takeAll(true) end, button, -2)
      button = _createButton("TA", toolbar, function() takeAll() end, button, -2)      
      
      local editBox = CreateFrame("EditBox", "BulkMailInboxSearchFilterEditBox", toolbar, "InputBoxTemplate")
      editBox:SetWidth(100)
      editBox:SetHeight(30);
      editBox:SetScript("OnTextChanged",
			function()
			   filterText = editBox:GetText():lower()
			   wipe(markTable)
			   mod:RefreshInboxGUI()
			end)
      
      editBox:SetScript("OnEscapePressed", editBox.ClearFocus)
      editBox:SetScript("OnEnterPressed", editBox.ClearFocus)
      
      editBox:SetAutoFocus(false)
      editBox:SetPoint("RIGHT", button, "LEFT", -10, 0);

      local text = toolbar:CreateFontString(nil, nil, "GameFontNormal")
      text:SetTextColor(1,1,1,1)
      text:SetText(color(L["Search"]..": ", "ffd200"))
      text:SetPoint("RIGHT", editBox, "LEFT", -5, 0)

      local text = toolbar:CreateFontString(nil, nil, "GameTooltipHeaderText")
      text:SetTextColor(1,1,1,1)
      text:SetText(color(L["Bulk Mail Inbox"], "ffd200"))
      text:SetPoint("LEFT", toolbar, "LEFT", 5, 0)
       

      local backdrop = GameTooltip:GetBackdrop()
      

      toolbar:SetBackdrop(backdrop)
   
      if backdrop then
	 toolbar:SetBackdropColor(GameTooltip:GetBackdropColor())
	 toolbar:SetBackdropBorderColor(GameTooltip:GetBackdropBorderColor())
      end

      toolbar:EnableMouse(true)
      toolbar:RegisterForDrag("LeftButton")
      toolbar:SetMovable(true)
      
      mod._toolbar = toolbar
   end
   
   toolbar:SetScript("OnDragStart", function() tooltip:StartMoving() end)
   toolbar:SetScript("OnDragStop", function() tooltip.moved = true tooltip:StopMovingOrSizing() end)
   
   toolbar:ClearAllPoints()
   toolbar:SetParent(tooltip)
   
   toolbar:SetPoint("BOTTOMLEFT", tooltip, "TOPLEFT", 0, -4)
   toolbar:SetPoint("BOTTOMRIGHT", tooltip, "TOPRIGHT", 0, -4)

   toolbar:Show()
end
      

-- For pagination
local MAX_ROWS = 80
local startPage = 0


-- This adds the header info, and next prev buttons if needed
local function _addHeaderAndNavigation(tooltip, firstRow, lastRow)
   local y = tooltip:AddLine();
   if firstRow and lastRow then
      tooltip:SetCell(y, 1, color(fmt(L["Inbox Items (%d mails, %s)"], GetInboxNumItems(), abacus:FormatMoneyShort(inboxCash)), "ffd200"), tooltip:GetFont(), "LEFT", 4)
      tooltip:SetCell(y, 5, color(fmt(L["Item %d-%d of %d"], firstRow, lastRow, numInboxItems), "ffd200"), tooltip:GetFont(), "RIGHT", 3)
	 
      y = tooltip:AddLine();
      if startPage > 0 then
	 tooltip:SetCell(y, 1,  color("<- "..L["Previous Page"], "ffd200"), tooltip:GetFont(), "LEFT", 2)
	 tooltip:SetCellScript(y, 1, "OnMouseUp", function() startPage = startPage - 1 mod:ShowInboxGUI() end)
      end
      
      if lastRow < numInboxItems then
	 tooltip:SetCell(y, 5,  color(L["Next Page"].." ->", "ffd200"), tooltip:GetFont(), "RIGHT", 3)
	 tooltip:SetCellScript(y, 5, "OnMouseUp", function() startPage = startPage + 1 mod:ShowInboxGUI() end)
      end
   else
      tooltip:SetCell(y, 1, color(fmt(L["Inbox Items (%d mails, %d items, %s)"], GetInboxNumItems(), numInboxItems, abacus:FormatMoneyShort(inboxCash)), "ffd200"), tooltip:GetFont(), "LEFT", 7)
   end
   
   tooltip:AddLine(" ")
   local sel = function(str, col)
		  return color(str, col == mod.db.char.sortField and "ffff7f" or "ffffff")
	       end
   y = tooltip:AddLine("", sel(L["Items (Inbox click actions apply)"], 1), sel(L["Qty."], 2), sel(L["Returnable"], 3), sel(L["Sender"], 4), sel(L["TTL"], 5), sel(L["Mail #"], 6))
   local setSortFieldFunc = function(obj, field)
			       if mod.db.char.sortField == field then
				  mod.db.char.sortReversed = not mod.db.char.sortReversed and true or nil
			       else
				  mod.db.char.sortReversed = nil
				  mod.db.char.sortField = field
			       end
			       self:RefreshInboxGUI()
			    end
   for i = 1,6 do 
      tooltip:SetCellScript(y, i+1, "OnMouseUp", setSortFieldFunc, i)
   end
   tooltip:AddSeparator(2)
end

local function _adjustSizeAndPosition(tooltip)
   tooltip:UpdateScrolling(UIParent:GetHeight() / tooltip:GetScale() * 0.7)
   tooltip:SetClampedToScreen(true)

   if not tooltip.moved then
      -- only adjust point if user hasn't moved manually
      tooltip:ClearAllPoints()
      tooltip:SetPoint("TOPLEFT", MailFrame, "TOPRIGHT", -5, -40)
   end
end

-- Update the clickability state of all buttons that toggle state
local function _updateButtonStates(tooltip)
   local hasMarked = next(markTable)
   local markColor = hasMarked and "ffd200" or "7f7f7f"
   
   if hasMarked then
      mod.buttons.CS:Enable()
      mod.buttons.TS:Enable()
   else
      mod.buttons.CS:Disable()
      mod.buttons.TS:Disable()
   end
   if inboxCash > 0 then
      mod.buttons.TC:Enable();
   else
      mod.buttons.TC:Disable();
   end
  
   mod.cells.takeSelected = _addColspanCell(tooltip, color(L["Take Selected"], markColor), 2, hasMarked and function() takeAll(false, true) end, nil, mod.cells.takeSelected)
   mod.cells.clearSelected = _addColspanCell(tooltip, color(L["Clear Selected"], markColor), 2, hasMarked and function() wipe(markTable) self:RefreshInboxGUI() end, nil, mod.cells.clearSelected)

   -- Re-set this or the tooltip freaks out.
   _adjustSizeAndPosition(tooltip)
end

      
function mod:ShowInboxGUI()
   if not mod.db.char.inboxUI then return end
   if not inboxCache or not next(inboxCache) then
      inboxCacheBuild()
   end

   local tooltip = mod.inboxGUI

   if not tooltip then
      tooltip = QTIP:Acquire("BulkMailInboxGUI")
      tooltip:EnableMouse(true)
      tooltip:SetScript("OnDragStart", tooltip.StartMoving)
      tooltip:SetScript("OnDragStop", function() tooltip.moved = true tooltip:StopMovingOrSizing() end)
      tooltip:RegisterForDrag("LeftButton")
      tooltip:SetMovable(true)
      tooltip:SetColumnLayout(7, "LEFT", "LEFT", "CENTER", "CENTER", "CENTER", "CENTER", "CENTER")
      mod.inboxGUI = tooltip
      startPage = 0
      tooltip:SetPoint("TOPLEFT", MailFrame, "TOPRIGHT", -5, -40)
   else
      tooltip:Clear()      
   end
   local y
   _createOrAttachSearchBar(tooltip)
   
   local markedColor = function(str, marked, col)
			  return color(str, col == mod.db.char.sortField and "ffff7f" or (marked and "ffffff" or "ffd200"))
		       end
   if inboxCache and #inboxCache > 0 then
      local firstRow, lastRow
      if numInboxItems > MAX_ROWS then
	 firstRow = MAX_ROWS * startPage + 1
	 while firstRow > numInboxItems and startPage >= 0 do
	    startPage = startPage - 1
	    firstRow = MAX_ROWS * startPage
	 end
	 lastRow = math.min(firstRow+MAX_ROWS, numInboxItems)

	 y = tooltip:AddLine()
	 _addHeaderAndNavigation(tooltip, firstRow, lastRow)
      else
	 startPage = 0
	 firstRow = 1
	 lastRow = numInboxItems
	 _addHeaderAndNavigation(tooltip)
      end
      for i = firstRow, lastRow do
	 local info = inboxCache[i]
	 local isMarked = markTable[info.bmid]
	 local itemText = info.itemLink or L["Cash"]
	 if info.texture then
	    itemText = fmt("|T%s:18|t%s", info.texture, itemText)
	 end
	 y = tooltip:AddLine("", 
			     itemText,
			     markedColor(info.money and abacus:FormatMoneyFull(info.money) or info.qty, isMarked, 2), 
			     markedColor(info.returnable and L["Yes"] or L["No"], isMarked, 3), 
			     markedColor(info.sender, isMarked, 4), 
			     markedColor(fmt("%0.1f", info.daysLeft), isMarked, 5), 
			     markedColor(info.index, isMarked, 6))

	 tooltip:SetCell(y, 1, isMarked and [[|TInterface\Buttons\UI-CheckBox-Check:18|t]] or " ", nil,  "RIGHT", 1, nil, 0, 0, 15, 15)
	 
	 tooltip:SetLineScript(y, "OnMouseUp", function(frame, line)
						  if not IsModifierKeyDown() then
						     if info.bmid then 
							markTable[info.bmid] = not markTable[info.bmid] and true or nil
							tooltip:SetCell(line, 1, markTable[info.bmid] and [[|TInterface\Buttons\UI-CheckBox-Check:18|t]] or " ", nil,  "RIGHT", 1, nil, 0, 0, 15, 15)
							_updateButtonStates(tooltip)
						     end
						  elseif info.index and info.attachment then
						     mod:InboxFrame_OnClick(nil, info.index, info.attachment)
						  elseif info.index and info.money then
						     mod:InboxFrame_OnClick(nil, info.index, info.money)
						  end
					       end, y)
	 tooltip:SetLineScript(y, "OnEnter", _onEnterFunc, info)
	 tooltip:SetLineScript(y, "OnLeave", _onLeaveFunc, info)
      end
   else
      _addHeaderAndNavigation(tooltip)
      tooltip:AddLine(" ", L["No items"])
   end
   tooltip:AddLine(" ")

   local cells = mod.cells or {}
   wipe(cells)
   mod.cells = cells

   _addColspanCell(tooltip, color(L["Take All"], "ffd200"), 7, function() takeAll() end)
   _addColspanCell(tooltip, color(L["Take Cash"], inboxCash > 0 and "ffd200" or "7f7f7f"), 7, inboxCash > 0 and function() takeAll(true) end)
   _updateButtonStates(tooltip)
   _addColspanCell(tooltip, color(L["Close"], "ffd200"), 7, mod.HideInboxGUI, mod)


   tooltip:SetFrameStrata("FULLSCREEN")
   -- set max height to be 80% of the screen height
   _adjustSizeAndPosition(tooltip)
   
   tooltip:Show()
end

-- Convenience function for registering options tables
function mod:OptReg(optname, tbl, cmd)
   local regtable
   local configPanes = self.configPanes or {}
   self.configPanes = configPanes
   AC:RegisterOptionsTable(optname, tbl, cmd)
   if not BulkMail then
      -- Only add it to the UI if it's not already added to Bulk Mail
      regtable = ACD:AddToBlizOptions(optname, L["Bulk Mail Inbox"])
   end
   configPanes[#configPanes+1] = optname
   return regtable
end
function mod:OpenConfigMenu(parentframe)
   -- create the menu   
   local frame = LD:OpenAce3Menu(mod.opts)

   -- Anchor the menu to the mouse
   frame:SetPoint("TOPLEFT", parentframe, "BOTTOMLEFT", 0, 0)
   frame:SetFrameLevel(parentframe:GetFrameLevel()+100)
end

function mod:ToggleConfigDialog()
   if mod._mainConfig then
      InterfaceOptionsFrame_OpenToCategory(mod._mainConfig)
   end
end
