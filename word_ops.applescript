-- Word 桥接：所有对 Word 的操作都在这里。由 server.py 通过 osascript 调用。
-- 输出格式：字段用 ASCII 31 分隔、记录用 ASCII 30 分隔，第一个字段是 OK 或 ERR。
-- 约定：只做带修订痕迹的修改；从不保存、关闭、接受或拒绝修订。

property fieldSep : character id 31
property recSep : character id 30

on joinList(lst, sep)
	set AppleScript's text item delimiters to sep
	set s to lst as text
	set AppleScript's text item delimiters to ""
	return s
end joinList

-- 精确（区分大小写等）按 findText 切分 t
on splitExact(t, findText)
	considering case, diacriticals, hyphens, punctuation and white space
		set AppleScript's text item delimiters to findText
		set parts to text items of t
		set AppleScript's text item delimiters to ""
	end considering
	return parts
end splitExact

on textsEqual(a, b)
	considering case, diacriticals, hyphens, punctuation and white space
		return a is b
	end considering
end textsEqual

-- [a,b) 区间内的“可见”字符数（不含已被修订删除的文字、域代码等）
on vis(d, a, b)
	if b ≤ a then return 0
	tell application "Microsoft Word"
		set r to create range d start a end b
		set s to content of r
	end tell
	return (count of s)
end vis

-- 把“可见文本中的偏移 k”换算成 Word 文档位置：
-- 返回最大的 p，使 [0,p) 的可见字符数 ≤ k
on posForVisible(d, k, docEnd, excess)
	if k ≤ 0 then
		set base to 0
		set target to 0
	else
		set base to k
		if base > docEnd then set base to docEnd
		set target to k - (my vis(d, 0, base))
	end if
	set lo to base
	set hi to base + excess + 1
	if hi > docEnd then set hi to docEnd
	repeat while lo < hi
		set mid to (lo + hi + 1) div 2
		if (my vis(d, base, mid)) ≤ target then
			set lo to mid
		else
			set hi to mid - 1
		end if
	end repeat
	return lo
end posForVisible

-- 从 p 起最小的 q，使 [p,q) 的可见字符数 ≥ need
on endForVisible(d, p, need, docEnd)
	if need ≤ 0 then return p
	set lo to p
	set hi to docEnd
	-- 先指数扩张，避免每次都取到文末
	set span to need
	repeat
		set cand to p + span
		if cand ≥ docEnd then
			set hi to docEnd
			exit repeat
		end if
		if (my vis(d, p, cand)) ≥ need then
			set hi to cand
			exit repeat
		end if
		set span to span * 2
	end repeat
	repeat while lo < hi
		set mid to (lo + hi) div 2
		if (my vis(d, p, mid)) ≥ need then
			set hi to mid
		else
			set lo to mid + 1
		end if
	end repeat
	return lo
end endForVisible

-- 在文档里定位 findText 的第 occ 次出现（occ=0 表示必须唯一）。
-- 返回 {状态, 出现次数, 可见偏移 k, 文档全文 t, docEnd, excess}
on locateVisible(d, findText, occ)
	tell application "Microsoft Word"
		set t to content of text object of d
		set docEnd to end of content of text object of d
	end tell
	set parts to my splitExact(t, findText)
	set n to (count of parts) - 1
	if n < 1 then return {"notfound", 0, 0, t, docEnd, 0}
	if occ is 0 then
		if n > 1 then return {"ambiguous", n, 0, t, docEnd, 0}
		set occ to 1
	end if
	if occ > n then return {"badocc", n, 0, t, docEnd, 0}
	set k to 0
	set fl to count of findText
	repeat with i from 1 to occ
		set k to k + (count of (item i of parts))
		if i < occ then set k to k + fl
	end repeat
	set excess to docEnd - (count of t)
	if excess < 0 then set excess to 0
	return {"ok", n, k, t, docEnd, excess}
end locateVisible

on run argv
	set op to item 1 of argv
	try
		with timeout of 120 seconds
			if op is "list" then
				tell application "Microsoft Word"
					if not running then return "OK" & my fieldSep
					set nms to name of every document
				end tell
				return "OK" & my fieldSep & my joinList(nms, my recSep)
			end if

			set docName to item 2 of argv
			tell application "Microsoft Word"
				if not (exists document docName) then return "ERR" & my fieldSep & "not_open"
				set d to document docName
			end tell

			if op is "read" then
				tell application "Microsoft Word"
					set t to content of text object of d
					set tr to track revisions of d
					set nr to count of revisions of d
					set nc to count of Word comments of d
					set sv to saved of d
				end tell
				return "OK" & my fieldSep & (tr as text) & my fieldSep & nr & my fieldSep & nc & my fieldSep & (sv as text) & my fieldSep & t

			else if op is "edit" then
				-- argv: edit doc findText prefix middleOld middleNew occurrence author
				set findText to item 3 of argv
				set prefixText to item 4 of argv
				set middleOld to item 5 of argv
				set middleNew to item 6 of argv
				set occ to (item 7 of argv) as integer
				set authorName to item 8 of argv
				set {locStatus, n, k, t, docEnd, excess} to my locateVisible(d, findText, occ)
				if locStatus is not "ok" then return "ERR" & my fieldSep & locStatus & my fieldSep & n
				set kStart to k + (count of prefixText)
				set p to my posForVisible(d, kStart, docEnd, excess)
				set q to my endForVisible(d, p, (count of middleOld), docEnd)
				tell application "Microsoft Word"
					if q > p then
						set got to content of (create range d start p end q)
					else
						set got to ""
					end if
				end tell
				if not (my textsEqual(got, middleOld)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & got
				tell application "Microsoft Word"
					set oldName to user name
					set track revisions of d to true
					try
						if authorName is not "" then set user name to authorName
						set r to create range d start p end q
						set content of r to middleNew
						set user name to oldName
					on error errMsg
						set user name to oldName
						return "ERR" & my fieldSep & "write_failed" & my fieldSep & errMsg
					end try
					set nr to count of revisions of d
				end tell
				return "OK" & my fieldSep & n & my fieldSep & p & my fieldSep & q & my fieldSep & nr

			else if op is "comment" then
				-- argv: comment doc anchorText occurrence commentText author
				set findText to item 3 of argv
				set occ to (item 4 of argv) as integer
				set commentText to item 5 of argv
				set authorName to item 6 of argv
				set {locStatus, n, k, t, docEnd, excess} to my locateVisible(d, findText, occ)
				if locStatus is not "ok" then return "ERR" & my fieldSep & locStatus & my fieldSep & n
				set p to my posForVisible(d, k, docEnd, excess)
				set q to my endForVisible(d, p, (count of findText), docEnd)
				tell application "Microsoft Word"
					set got to content of (create range d start p end q)
				end tell
				if not (my textsEqual(got, findText)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & got
				tell application "Microsoft Word"
					set oldName to user name
					try
						if authorName is not "" then set user name to authorName
						set r to create range d start p end q
						set c to make new Word comment at d with properties {scope:r, comment text:commentText}
						set user name to oldName
					on error errMsg
						set user name to oldName
						return "ERR" & my fieldSep & "write_failed" & my fieldSep & errMsg
					end try
					set nc to count of Word comments of d
				end tell
				return "OK" & my fieldSep & n & my fieldSep & nc

			else if op is "revisions" then
				set maxN to (item 3 of argv) as integer
				set revLines to {}
				tell application "Microsoft Word"
					set nr to count of revisions of d
					set upTo to nr
					if upTo > maxN then set upTo to maxN
					repeat with i from 1 to upTo
						set rv to revision i of d
						set end of revLines to ((revision type of rv) as text) & my fieldSep & (author of rv) & my fieldSep & (content of text object of rv)
					end repeat
					set nc to count of Word comments of d
					set comLines to {}
					set upC to nc
					if upC > maxN then set upC to maxN
					repeat with i from 1 to upC
						set c to Word comment i of d
						set end of comLines to (author of c) & my fieldSep & (content of scope of c) & my fieldSep & (content of comment text of c)
					end repeat
				end tell
				return "OK" & my fieldSep & nr & my fieldSep & nc & my recSep & my joinList(revLines, my recSep) & (character id 29) & my joinList(comLines, my recSep)
			end if
			return "ERR" & my fieldSep & "unknown_op"
		end timeout
	on error errMsg number errNum
		return "ERR" & my fieldSep & "exception" & my fieldSep & errMsg & " (" & errNum & ")"
	end try
end run
