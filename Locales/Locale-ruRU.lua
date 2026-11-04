local L = LibStub("AceLocale-3.0"):NewLocale("BulkMailInbox", "ruRU")
if not L then return end

-- To help localize BulkMailInbox please enter phrase translations on the
-- following URL:
-- http://www.wowace.com/projects/bulkmail-inbox/localization/
-- This file should not be edited manually!

--@localization(locale="ruRU", format="lua_additive_table", handle-subnamespaces="handle-subnamespaces")@

-- These phrases used to be single strings with a newline in the key. The previous
-- Russian translation was one paragraph; it is split on the same sentence boundary
-- so the tooltip text stays intact when the lines are joined again.
L["Filter the inbox display to items matching the term entered here."] = "Фильтрация списка входящих писем по введенному слову."
L["Take All and Take Cash actions are limited to items matching the inbox filter."] = "Действия «Забрать всё» и «Забрать деньги» будут применены только к отфильтрованным предметам."
L["Take all items from the mailbox. If the search filter is used,"] = "Забрать все предметы из почтового ящика."
L["only items matching the search term will be taken."] = "Если используется фильтр поиска, будут взяты только те предметы, которые соответствуют условию поиска."
L["Take all money from the mailbox. If the search filter is used,"] = "Забрать все деньги из почтового ящика."
L["money will only be taken from mails which the search term."] = "Если используется фильтр поиска, деньги будут взяты только из тех писем, которые соответствуют условию поиска."
