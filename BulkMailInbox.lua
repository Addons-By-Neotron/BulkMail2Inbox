BulkMailInbox = LibStub("AceAddon-3.0"):NewAddon("BulkMailInbox", "AceConsole-3.0", "AceEvent-3.0", "AceTimer-3.0", "AceHook-3.0")

local mod, self, BulkMailInbox = BulkMailInbox, BulkMailInbox, BulkMailInbox

local VERSION = " @project-version@"

local LibStub = LibStub

local L        = LibStub("AceLocale-3.0"):GetLocale("BulkMailInbox", false)

BulkMailInbox.L = L

local media  = LibStub("LibSharedMedia-3.0")
local abacus = LibStub("LibAbacus-3.0")
local QTIP   = LibStub("LibQTip-1.0")
local AC     = LibStub("AceConfig-3.0")
local ACD    = LibStub("AceConfigDialog-3.0")
local DB     = LibStub("AceDB-3.0")
local LDB    = LibStub("LibDataBroker-1.1", true)
local LD     = LibStub("LibDropdown-1.0")


local _G = _G
local fmt = string.format
local lower = string.lower

local sortFields, markTable  -- tables
local ibIndex, ibAttachIndex, numInboxItems, inboxCash, cleanPass, cashOnly, markOnly, takeAllInProgress, invFull, filterText -- variables
local spinnerText = { "Working   ", "Working.  ", "Working.. ", "Working..." }

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
local function _sortInboxFunc(itemA,itemB)
    local sf = sortFields[BulkMailInbox.db.char.sortField]
    if itemA and itemB then
        local a, b = itemA[sf], itemB[sf]
        if sf == 'itemLink' then
            a = a and GetItemInfo(a) or a
            b = b and GetItemInfo(b) or b
        elseif sf == 'qty' then
            a = a or itemA["money"]
            b = b or itemB["money"]
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
            if subject:find(pattern) ~= nil then
                return true
            end
        end
    end
end

local function _matchesFilter(text)
    if not filterText or filterText:len() == 0 then
        return true
    end
    return text:lower():find(filterText, 1, true) ~= nil
end

local function inboxCacheBuild()
    local start = GetTime()
    for k in ipairs(inboxCache) do inboxCache[k] = del(inboxCache[k]) end
    inboxCash, numInboxItems = 0, 0
    for i = 1, GetInboxNumItems() do
        local _, _, sender, subject, money, cod, daysLeft, numItems, _, wasReturned, _, canReply, isGM = GetInboxHeaderInfo(i)
        if money > 0 then
            local _, itemName = GetInboxInvoiceInfo(i)
            local title = itemName and ITEM_SOLD_COLON..' '..itemName or L["Cash"]
            if _matchesFilter(title) then
                -- Contributed by Scott Centoni
                table.insert(inboxCache, newHash(
                        'index', i, 'sender', sender, 'bmid', daysLeft..subject..0, 'returnable', not wasReturned, 'cod', cod,
                        'daysLeft', daysLeft, 'itemLink', title, 'money', money, 'texture', "Interface\\Icons\\INV_Misc_Coin_01"
                ))
                inboxCash = inboxCash + money
            end
        end
        if numItems then
            local canReturnItem = not wasReturned
            if isGM or not canReply or _isAHSentMail(subject) then
                canReturnItem = false
            end
            for j=1, ATTACHMENTS_MAX_RECEIVE do
                if GetInboxItem(i,j) and _matchesFilter(GetInboxItem(i, j)) then
                    table.insert(inboxCache, newHash(
                            'index', i, 'attachment', j, 'sender', sender, 'bmid', daysLeft..subject..j, 'returnable', canReturnItem, 'cod', cod,
                            'daysLeft', daysLeft, 'itemLink', GetInboxItemLink(i,j), 'qty', select(4, GetInboxItem(i,j)), 'texture', (select(3, GetInboxItem(i,j)))
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
    return fmt("|cff%s%s|r", color, text or "")
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
            takeStackable = true,
            sortField = 1,
        },
        profile = {
            disableTooltips = false,
            scale = 1.0,
            font = "Friz Quadrata TT",
            fontSize = 12,
            pageSize = 50
        }
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
            gui = {
                name = L["Show Inbox GUI"], type = 'toggle',
                desc = L["Show the Inbox Items GUI"],
                get = function() return self.db.char.inboxUI end,
                set = function(args,v) self.db.char.inboxUI = v self:RefreshInboxGUI() end,
            },
            takeStackable = {
                name = L["Always Take Stackable Items"], type = 'toggle',
                desc = L["Continue taking partial stacks of stackable items even if the mailbox is full."],
                get = function() return self.db.char.takeStackable end,
                set = function(args,v) self.db.char.takeStackable = v end,
            },
            disableTooltips = {
                name = L["Disable Tooltips"], type = 'toggle',
                desc = L["Disable the help tooltips for the toolbar buttons."],
                get = function() return self.db.profile.disableTooltips end,
                set = function(args,v) self.db.profile.disableTooltips = v end,
            },
            scale = {
                type = "range",
                name = L["GUI Scale"],
                desc = L["Set the window scale of the Inbox GUI."],
                min = 0.3, max = 3.0, step = 0.1,
                set = function(_,val) mod.db.profile.scale = val mod:RefreshInboxGUI(true) end,
                get = function() return mod.db.profile.scale end,
                order = 500,
            },
            pageSize = {
                type = "range",
                name = L["Page Size"],
                desc = L["Maximum number of items to display per page."],
                min = 10, max = 100, step = 1,
                set = function(_,val) mod.db.profile.pageSize = val mod:RefreshInboxGUI(true) end,
                get = function() return mod.db.profile.pageSize end,
                order = 500,
            },
            font = {
                type = "select",
                dialogControl = "LSM30_Font",
                name = L["Font"],
                desc = L["Font used in the inbox list"],
                values = AceGUIWidgetLSMlists.font,
                set = function(_,key) mod.db.profile.font = key  mod:RefreshInboxGUI() end,
                get = function() return mod.db.profile.font end,
                order = 1000,
            },
            fontsize = {
                type = "range",
                name = L["Font size"],
                min = 6, max = 30, step = 1,
                set = function(_,val) mod.db.profile.fontSize = val mod:RefreshInboxGUI() end,
                get = function() return mod.db.profile.fontSize end,
                order = 2000,
            },

        },
    }

    -- set up LDB, but only if the user doesn't have Bulk Mail already
    if LDB and not BulkMail then
        self.ldb =
        LDB:NewDataObject("BulkMailInbox",
                {
                    type =  "data source",
                    label = L["Bulk Mail Inbox"]..VERSION,
                    icon = [[Interface\Addons\BulkMail2Inbox\icon]],
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

    if BulkMail and BulkMail.opts then
        BulkMail.opts.args.inbox = { type = "group",
                                     handler = mod,
                                     name = L["Inbox"],
                                     desc = L["Bulk Mail Inbox Options"],
                                     args = BulkMailInbox.opts.args
        }
    end
end

function mod:OnEnable()
    self:RegisterEvent('MAIL_SHOW')
    self:RegisterEvent('MAIL_CLOSED')
    self:RegisterEvent('PLAYER_ENTERING_WORLD')
    self:RegisterEvent('UI_ERROR_MESSAGE')
    self:RegisterEvent('MAIL_INBOX_UPDATE')
    if not _G.GetContainerItemInfo then
        self:RegisterEvent('PLAYER_INTERACTION_MANAGER_FRAME_HIDE')
    end
    -- Handle being LoD loaded while at the mailbox
    if MailFrame:IsVisible() then
        self:MAIL_SHOW()
    end
end

function mod:OnDisable()
    self:UnregisterAllEvents()
end

------------------------------------------------------------------------------
-- Events
------------------------------------------------------------------------------
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

function mod:PLAYER_INTERACTION_MANAGER_FRAME_HIDE(_, type)
    if type == Enum.PlayerInteractionType.MailInfo then
        mod:MAIL_CLOSED()
    end
end

function mod:MAIL_CLOSED()
    takeAllInProgress = false
    self:HideInboxGUI()
    GameTooltip:Hide()
    self:UnhookAll()
end

BulkMailInbox.PLAYER_ENTERING_WORLD = BulkMailInbox.MAIL_CLOSED  -- MAIL_CLOSED doesn't get called if, for example, the player accepts a port with the mail window open

function mod:UI_ERROR_MESSAGE(event, type, msg)  -- prevent infinite loop when inventory is full
    if msg == ERR_INV_FULL then
        invFull = true
    end
end

-- Take next inbox item or money skip past CoD items and letters.
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

local _fetchCount = 0
local _lastCount = -1
local function _updateSpinner()
    local spinner = mod._toolbar and mod._toolbar.spinner
    if not spinner or (takeAllInProgress and _fetchCount == _lastCount) then return end
    _lastCount = _fetchCount
    local isVisible = mod.buttons.Cancel:IsVisible()
    if takeAllInProgress then
        spinner:SetText(spinnerText[1+math.fmod(_fetchCount, #spinnerText)]);
        if not isVisible then
            mod.buttons.Cancel:Show()
        end
    elseif isVisible then
        mod.buttons.Cancel:Hide()
        spinner:SetText("")
    end
end

function mod:TakeNextItemFromMailbox()
    _updateSpinner()
    if not takeAllInProgress then
        return
    end

    local numMails = GetInboxNumItems()
    cashOnly = cashOnly or (invFull and not mod.db.char.takeStackable)

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
    local sender, subject, money, cod, daysLeft, item, _, _, text, _, isGM = select(3, GetInboxHeaderInfo(curIndex))

    if subject then
        prevSubject = subject
    else
        subject = prevSubject
    end

    if curAttachIndex == ATTACHMENTS_MAX_RECEIVE then
        ibIndex = ibIndex - 1
        ibAttachIndex = 0
    else
        ibAttachIndex = ibAttachIndex + 1
    end
    local itemName, _, _, itemCount = GetInboxItem(curIndex, curAttachIndex)
    local markKey = daysLeft..subject..curAttachIndex

    if (sender == "The Postmaster" or sender == "Thaumaturge Vashreen") and not itemName and money == 0 and not item then
        DeleteInboxItem(curIndex)
        self:SmartScheduleTimer('BMI_RefreshInboxGUI', false, "RefreshInboxGUI", 1)
        self:SmartScheduleTimer('BMI_TakeNextItem', true, "TakeNextItemFromMailbox", 0.4)
        return
    end

    if curAttachIndex > 0 and not itemName or markOnly and not markTable[markKey] or itemName and not _matchesFilter(itemName)
    then
        return self:TakeNextItemFromMailbox()
    end
    local actionTaken
    if not string.find(subject, "Sale Pending") then
        if curAttachIndex == 0 and money > 0 then
            local _, itemName = GetInboxInvoiceInfo(curIndex)
            local title = itemName and ITEM_SOLD_COLON..' '..itemName or L["Cash"]
            if _matchesFilter(title) then
                cleanPass = false
                actionTaken = true
                TakeInboxMoney(curIndex)
            end
        elseif not cashOnly and cod == 0 then
            cleanPass = invFull -- this ensures we'll die properly after a full mailbox iteration
            local inboxitem = GetInboxItemLink(curIndex,curAttachIndex)
            if inboxitem and
                    (not invFull or -- inventory not full
                            (mod.db.char.takeStackable and -- or continue taking stackable items even if full
                                    itemCount < select(8, GetItemInfo(inboxitem)))) then
                TakeInboxItem(curIndex, curAttachIndex)
                markTable[markKey] = nil
                actionTaken = true
            end
        end
    end

    if actionTaken then
        -- Since we did something, we'll add a delay to prevent erroring out
        self:SmartScheduleTimer('BMI_RefreshInboxGUI', false, "RefreshInboxGUI", 1)
        self:SmartScheduleTimer('BMI_TakeNextItem', true, "TakeNextItemFromMailbox", 0.4)
        _fetchCount = _fetchCount + 1
    else
        -- We didn't take any items so let's move on
        self:TakeNextItemFromMailbox()
    end
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
-- For pagination
local startPage = 0

local function _closeHelpTooltip(parentFrame)
    if mod.helpTooltip and mod.helpTooltip.owner == parentFrame then
        mod.helpTooltip.owner = nil
        QTIP:Release(mod.helpTooltip)
        mod.helpTooltip = nil
    end
end

local function _openHelpTooltip(parentFrame, header, text)
    if self.db.profile.disableTooltips then return end
    local tooltip = mod.helpTooltip or QTIP:Acquire("BulkMailInboxHelpTooltip")
    mod.helpTooltip = tooltip
    tooltip:Clear()

    tooltip.owner = parentFrame
    tooltip:SetColumnLayout(1, "LEFT")
    tooltip:AddHeader(header)
    tooltip:AddLine(color(text, "ffd200"))
    tooltip:SmartAnchorTo(parentFrame)
    tooltip:SetClampedToScreen(true)
    tooltip:Show()
end

local function _addTooltipToFrame(frame, header, text)
    frame:SetScript("OnEnter", function(self) _openHelpTooltip(self, header, text) end)
    frame:SetScript("OnLeave", _closeHelpTooltip)
end


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
    mod:SmartCancelTimer('BMI_takeAll')
    mod:SmartCancelTimer('BMI_TakeNextItem')
    mod:SmartCancelTimer('BMI_RefreshInboxGUI')

    if mod._toolbar then
        mod._toolbar:Hide()
        mod._toolbar:SetParent(nil)
    end

    local tooltip = mod.inboxGUI
    if tooltip then
        mod.inboxGUI = nil
        tooltip:EnableMouse(false)
        tooltip:SetScript("OnDragStart", nil)
        tooltip:SetScript("OnDragStop", nil)
        tooltip:SetMovable(false)
        tooltip:RegisterForDrag()
        tooltip:SetFrameStrata("TOOLTIP")
        tooltip.moved = nil
        tooltip:SetScale(GameTooltip:GetScale())
        QTIP:Release(tooltip)
    end
    mod._wantGui = nil
end

function mod:RefreshInboxGUI(resetMoved)
    _updateSpinner()
    mod:SmartCancelTimer('BMI_RefreshInboxGUI')
    if not mod.db.char.inboxUI then return end
    inboxCacheBuild()
    if mod.inboxGUI then
        if resetMoved then
            mod.inboxGUI.moved = nil
        end
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
                frame:Hide()
            end
        end
        GameTooltip.comparing = false
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

local function _createButton(title, parent, onclick, anchorTo, xoffset, tooltipHeader, tooltipText)
    local buttons = mod.buttons or {}
    mod.buttons = buttons

    local button = CreateFrame("Button", nil, parent, "UIPanelButtonTemplate")
    button:SetText(title)
    button:SetWidth(25)
    button:SetHeight(20)
    button:SetScript("OnClick", onclick)
    buttons[title] = button
    button:SetPoint("RIGHT", anchorTo, "LEFT", xoffset, 0)
    _addTooltipToFrame(button, tooltipHeader, tooltipText)
    return button
end

local function _createOrAttachSearchBar(tooltip)
    local toolbar = mod._toolbar
    if not toolbar then
        local template = (TooltipBackdropTemplateMixin and "TooltipBackdropTemplate") or (BackdropTemplateMixin and "BackdropTemplate")

        toolbar = CreateFrame("Frame", nil, UIParent, template)
        toolbar:SetHeight(49)

        local closeButton =  CreateFrame("Button", "BulkMailInboxToolbarCloseButton", toolbar, "UIPanelCloseButton")
        closeButton:SetPoint("TOPRIGHT", toolbar, "TOPRIGHT", 0, 0)
        closeButton:SetScript("OnClick", function() mod:HideInboxGUI() end)
        _addTooltipToFrame(closeButton, L["Close"], L["Close the window and stop taking items from the inbox."])

        local nextButton = CreateFrame("Button", nil, toolbar)
        nextButton:SetNormalTexture([[Interface\Buttons\UI-SpellbookIcon-NextPage-Up]])
        nextButton:SetPushedTexture([[Interface\Buttons\UI-SpellbookIcon-NextPage-Down]])
        nextButton:SetDisabledTexture([[Interface\Buttons\UI-SpellbookIcon-NextPage-Disabled]])
        nextButton:SetHighlightTexture([[Interface\Buttons\UI-Common-MouseHilight]], "ADD")
        nextButton:SetPoint("TOP", closeButton, "BOTTOM", 0, 9)
        nextButton:SetScript("OnClick", function() startPage = startPage + 1 mod:ShowInboxGUI() end)
        nextButton:SetWidth(25)
        nextButton:SetHeight(25)
        _addTooltipToFrame(nextButton, L["Next Page"], L["Go to the next page of items."])

        local prevButton = CreateFrame("Button", nil, toolbar)
        prevButton:SetNormalTexture([[Interface\Buttons\UI-SpellbookIcon-PrevPage-Up]])
        prevButton:SetPushedTexture([[Interface\Buttons\UI-SpellbookIcon-PrevPage-Down]])
        prevButton:SetDisabledTexture([[Interface\Buttons\UI-SpellbookIcon-PrevPage-Disabled]])
        prevButton:SetHighlightTexture([[Interface\Buttons\UI-Common-MouseHilight]], "ADD")
        prevButton:SetPoint("RIGHT", nextButton, "LEFT", 0, 0)
        prevButton:SetScript("OnClick", function() startPage = startPage - 1 mod:ShowInboxGUI() end)
        prevButton:SetWidth(25)
        prevButton:SetHeight(25)
        _addTooltipToFrame(prevButton, L["Previous Page"], L["Go to the previous page of items."])

        local pageText = toolbar:CreateFontString(nil, nil, "GameFontNormalSmall")
        pageText:SetTextColor(1,210/255.0,0,1)
        pageText:SetPoint("RIGHT", prevButton, "LEFT", 0, 0)
        toolbar.pageText = pageText

        local itemText = toolbar:CreateFontString(nil, nil, "GameFontNormalSmall")
        itemText:SetTextColor(1,210/255.0,0,1)
        itemText:SetPoint("TOPRIGHT", pageText, "TOPLEFT", 0, 0)
        itemText:SetPoint("BOTTOMRIGHT", pageText, "BOTTOMLEFT", 0, 0)
        itemText:SetPoint("LEFT", toolbar, "LEFT", 5, 0)
        itemText:SetJustifyH("LEFT")
        toolbar.itemText = itemText


        local button = _createButton("CS", toolbar, function() wipe(markTable) self:RefreshInboxGUI() end, closeButton, -2,
                L["Clear Selected"], L["Clear the list of selected items."])
        button = _createButton("TS", toolbar, function() takeAll(false, true) end, button, -2,
                L["Take Selected"], L["Take all selected items from the mailbox."])
        button = _createButton("TC", toolbar, function() takeAll(true) end, button, -2,
                L["Take Cash"], L["Take all money from the mailbox. If the search filter is used,\nmoney will only be taken from mails which the search term."])
        button = _createButton("TA", toolbar, function() takeAll() end, button, -2,
                L["Take All"], L["Take all items from the mailbox. If the search filter is used,\nonly items matching the search term will be taken."])

        mod.buttons.prev = prevButton
        mod.buttons.next = nextButton
        mod.buttons.close = closeButton

        local editBox = CreateFrame("EditBox", "BulkMailInboxSearchFilterEditBox", toolbar, "InputBoxTemplate")
        editBox:SetWidth(100)
        editBox:SetHeight(30)
        editBox:SetScript("OnTextChanged",
                function()
                    -- stop taking items when search terms change or we might
                    -- end up taking stuff we didn't mean to take
                    mod:SmartCancelTimer('BMI_takeAll')
                    takeAllInProgress = false
                    _updateSpinner()
                    filterText = editBox:GetText():lower()
                    wipe(markTable)
                    mod:RefreshInboxGUI()
                end)

        editBox:SetScript("OnEscapePressed", editBox.ClearFocus)
        editBox:SetScript("OnEnterPressed", editBox.ClearFocus)

        editBox:SetAutoFocus(false)
        editBox:SetPoint("RIGHT", button, "LEFT", -10, 0)
        _addTooltipToFrame(editBox, L["Search"], L["Filter the inbox display to items matching the term entered here.\nTake All and Take Cash actions are limited to items matching the inbox filter."])


        local text = toolbar:CreateFontString(nil, nil, "GameFontNormal")
        text:SetTextColor(1,210/255.0,0,1)
        text:SetText(L["Search"]..": ")
        text:SetPoint("RIGHT", editBox, "LEFT", -5, 0)



        local spinner = toolbar:CreateFontString(nil, nil, "GameFontNormal")
        spinner:SetTextColor(1,210/255.0,0,1)
        spinner:SetText("")
        spinner:SetPoint("TOPLEFT", text, "BOTTOMLEFT", 0, -10)
        spinner:SetJustifyH("RIGHT")
        toolbar.spinner = spinner

        local cancelButton =  CreateFrame("Button", "BulkMailInboxToolbarCancelButton", toolbar, "UIPanelCloseButton")
        cancelButton:SetPoint("RIGHT", spinner, "LEFT", 0, 0)
        cancelButton:SetScript("OnClick", function(self) takeAllInProgress = nil end)
        _addTooltipToFrame(cancelButton, L["Cancel"], L["Cancel taking items from the inbox."])
        mod.buttons.Cancel = cancelButton

        local titleText = toolbar:CreateFontString(nil, nil, "GameTooltipHeaderText")
        titleText:SetTextColor(1,210/255.0,0,1)
        titleText:SetText(L["Bulk Mail Inbox"])
        titleText:SetPoint("TOPRIGHT", text, "TOPLEFT", -5, 0)
        titleText:SetPoint("BOTTOMRIGHT", text, "BOTTOMLEFT", -5, 0)
        titleText:SetPoint("LEFT", toolbar, "LEFT", 5, 0)
        toolbar.titleText = titleText


        if TooltipBackdropTemplateMixin then
            tooltip.layoutType = GameTooltip.layoutType
            if GameTooltip.layoutType then
                tooltip.NineSlice:SetCenterColor(GameTooltip.NineSlice:GetCenterColor())
                tooltip.NineSlice:SetBorderColor(GameTooltip.NineSlice:GetBorderColor())
            end
        else
            local backdrop = GameTooltip:GetBackdrop()

            tooltip:SetBackdrop(backdrop)

            if backdrop then
                tooltip:SetBackdropColor(GameTooltip:GetBackdropColor())
                tooltip:SetBackdropBorderColor(GameTooltip:GetBackdropBorderColor())
            end
        end


        toolbar:EnableMouse(true)
        toolbar:RegisterForDrag("LeftButton")
        toolbar:SetMovable(true)
        mod._toolbar = toolbar
        mod._toolbarEditBox = editBox
    end

    toolbar:SetScript("OnDragStart", function() tooltip:StartMoving() end)
    toolbar:SetScript("OnDragStop", function() tooltip.moved = true tooltip:StopMovingOrSizing() end)

    toolbar:ClearAllPoints()
    toolbar:SetParent(tooltip)

    toolbar:SetPoint("BOTTOMLEFT", tooltip, "TOPLEFT", 0, -4)
    toolbar:SetPoint("BOTTOMRIGHT", tooltip, "TOPRIGHT", 0, -4)

    toolbar:Show()
end



-- This adds the header info, and next prev buttons if needed
local function _addHeaderAndNavigation(tooltip, totalRows, firstRow, lastRow)
    local y
    mod._toolbar.itemText:SetText(fmt(L["Inbox Items (%d mails, %s)"], GetInboxNumItems(), abacus:FormatMoneyShort(inboxCash)))
    if firstRow and lastRow then
        mod._toolbar.pageText:SetText(fmt(L["Item %d-%d of %d"], firstRow, lastRow, totalRows))

        if startPage > 0 then
            mod.buttons.prev:Enable()
        else
            mod.buttons.prev:Disable()
        end

        if lastRow < totalRows then
            mod.buttons.next:Enable()
        else
            mod.buttons.next:Disable()
        end
    else
        mod.buttons.next:Disable()
        mod.buttons.prev:Disable()
        if totalRows > 0 then
            mod._toolbar.pageText:SetText(fmt(L["Item %d-%d"], 1, #inboxCache))
        else
            mod._toolbar.pageText:SetText("")
        end
    end

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

function mod:AdjustSizeAndPosition(tooltip)

    local scale = mod.db.profile.scale

    tooltip:SetScale(scale)
    if not tooltip.moved then
        -- this is needed to get the correct height for some reason.
        tooltip:ClearAllPoints()
        tooltip:SetPoint("TOP", UIParent, "TOP", 0, 0)
    end
    local barHeight = mod._toolbar:GetHeight()*scale
    local uiHeight = UIParent:GetHeight()
    tooltip:UpdateScrolling((uiHeight-barHeight+10)/scale)

    -- Only adjust point if user hasn't moved manually. This puts it lined up with the mail window
    -- or in the middle of the screen it's too large to fit from the top of the mail window and down
    if not tooltip.moved and MailFrame ~= nil and MailFrame:GetTop() ~= nil then
        local tipHeight = tooltip:GetHeight() * scale
        tooltip:ClearAllPoints()
        -- Calculate a good offset
        local offx = math.min((uiHeight - tipHeight - barHeight)/2, uiHeight + 12 - MailFrame:GetTop()*MailFrame:GetScale())+barHeight
        tooltip:SetPoint("TOPLEFT", UIParent, "TOPLEFT", MailFrame:GetRight()*MailFrame:GetScale()/scale, -offx/scale)
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
        mod.buttons.TC:Enable()
    else
        mod.buttons.TC:Disable()
    end

    mod.cells.takeSelected = _addColspanCell(tooltip, color(L["Take Selected"], markColor), 2, hasMarked and function() takeAll(false, true) end, nil, mod.cells.takeSelected)
    mod.cells.clearSelected = _addColspanCell(tooltip, color(L["Clear Selected"], markColor), 2, hasMarked and function() wipe(markTable) self:RefreshInboxGUI() end, nil, mod.cells.clearSelected)

    -- Re-set this or the tooltip freaks out.
    mod:AdjustSizeAndPosition(tooltip)
end


function mod:ShowInboxGUI()
    if not mod.db.char.inboxUI then return end
    if not inboxCache or not next(inboxCache) then
        inboxCacheBuild()
    end

    local tooltip = mod.inboxGUI

    local refocus = false
    local cursorpos = 0
    if mod._toolbarEditBox and mod._toolbarEditBox:HasFocus() then
        refocus = true
        cursorpos = mod._toolbarEditBox:GetCursorPosition()
    end

    if not tooltip then
        tooltip = QTIP:Acquire("BulkMailInboxGUI")
        if tooltip.SetScrollStep then
            tooltip:SetScrollStep(100)
        end
        tooltip:EnableMouse(true)
        tooltip:SetScript("OnDragStart", tooltip.StartMoving)
        tooltip:SetScript("OnDragStop", function() tooltip.moved = true tooltip:StopMovingOrSizing() end)
        tooltip:RegisterForDrag("LeftButton")
        tooltip:SetMovable(true)
        tooltip:SetColumnLayout(7, "LEFT", "LEFT", "CENTER", "CENTER", "CENTER", "CENTER", "CENTER")
        mod.inboxGUI = tooltip
        startPage = 0
    else
        tooltip:Hide()
        tooltip:Clear()
    end

    local y

    local fontName = media:Fetch("font", mod.db.profile.font)

    local font = mod.font or CreateFont("BulkMailInboxFont")
    font:CopyFontObject(GameFontNormal)
    font:SetFont(fontName, mod.db.profile.fontSize, "")
    mod.font = font
    _createOrAttachSearchBar(tooltip)
    tooltip:SetFont(font)

    local markedColor = function(str, col)
        return color(str, col == mod.db.char.sortField and "ffff7f" or "ffd200")
    end
    local maxRows = mod.db.profile.pageSize
    if inboxCache and #inboxCache > 0 then
        local firstRow, lastRow
        local totalRows = #inboxCache
        if totalRows > maxRows then
            firstRow = maxRows * startPage + 1
            while firstRow > totalRows and startPage >= 0 do
                startPage = startPage - 1
                firstRow = maxRows * startPage
            end
            lastRow = math.min(firstRow+maxRows, totalRows)
            _addHeaderAndNavigation(tooltip, totalRows, firstRow, lastRow)
        else
            startPage = 0
            firstRow = 1
            lastRow = totalRows
            _addHeaderAndNavigation(tooltip, totalRows)
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
                    markedColor(info.money and abacus:FormatMoneyFull(info.money) or info.qty, 2),
                    markedColor(info.returnable and L["Yes"] or L["No"], 3),
                    markedColor(info.sender, 4),
                    markedColor(fmt("%0.1f", info.daysLeft), 5),
                    markedColor(info.index, 6))
            if isMarked then
                tooltip:SetLineColor(y, 1, 1, 1, 0.3)
            end
            tooltip:SetCell(y, 1, isMarked and [[|TInterface\Buttons\UI-CheckBox-Check:18|t]] or " ", nil,  "RIGHT", 1, nil, 0, 0, mod.db.profile.fontSize + 3, mod.db.profile.fontSize + 3)

            tooltip:SetLineScript(y, "OnMouseUp", function(frame, line)
                if not IsModifierKeyDown() then
                    if info.bmid then
                        markTable[info.bmid] = not markTable[info.bmid] and true or nil
                        tooltip:SetCell(line, 1, markTable[info.bmid]
                                and [[|TInterface\Buttons\UI-CheckBox-Check:18|t]] or " ", nil,  "RIGHT", 1, nil, 0, 0, 18, 15)
                        if markTable[info.bmid] then
                            tooltip:SetLineColor(line, 1, 1, 1, 0.3)
                        else
                            tooltip:SetLineColor(line, 1, 1, 1, 0.0)
                        end

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
        _addHeaderAndNavigation(tooltip, 0)
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
    mod:AdjustSizeAndPosition(tooltip)

    tooltip:Show()
    if refocus then
        mod._toolbarEditBox:SetFocus(true)
        mod._toolbarEditBox:HighlightText(0,0)
        mod._toolbarEditBox:SetCursorPosition(cursorpos)
    end
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
