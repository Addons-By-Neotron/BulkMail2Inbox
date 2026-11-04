local L = LibStub("AceLocale-3.0"):NewLocale("BulkMailInbox", "zhCN")
if not L then return end

-- To help localize BulkMailInbox please enter phrase translations on the
-- following URL:
-- http://www.wowace.com/projects/bulkmail-inbox/localization/
-- This file should not be edited manually!

--@localization(locale="zhCN", format="lua_additive_table", handle-subnamespaces="handle-subnamespaces")@

-- These phrases used to be single strings with a newline in the key. The newline
-- made the zhCN CurseForge export an unfinished Lua string. The two halves below
-- are the previous translation, split on that same newline.
L["Filter the inbox display to items matching the term entered here."] = "在框内输入关键字以显示相应物品."
L["Take All and Take Cash actions are limited to items matching the inbox filter."] = "取出所有及取出金币的功能也受搜索过滤的限制."
L["Take all items from the mailbox. If the search filter is used,"] = "取出邮箱中的所有物品."
L["only items matching the search term will be taken."] = "如果使用了搜索过滤,则只取出对应搜索关键字的物品."
L["Take all money from the mailbox. If the search filter is used,"] = "取出邮箱中所有含金币邮件."
L["money will only be taken from mails which the search term."] = "如果使用了搜索过滤,则只取出对应搜索关键字且包含金币的邮件."
