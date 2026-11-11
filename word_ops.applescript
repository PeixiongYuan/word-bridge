-- Word bridge: every operation on Word lives here. Called by server.py through osascript.
-- Output format: fields are separated by ASCII 31 and records by ASCII 30; the first field is OK or ERR.
-- Contract: only tracked edits are made; nothing is ever saved, closed, accepted or rejected.

property fieldSep : character id 31
property recSep : character id 30

on joinList(lst, sep)
	set AppleScript's text item delimiters to sep
	set s to lst as text
	set AppleScript's text item delimiters to ""
	return s
end joinList

-- Split t on findText exactly (case-sensitive and so on)
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

-- Number of "visible" characters in [a,b), not counting tracked-deleted text, field codes and the like.
-- Only reliable when a is a "clean" position: 0, or the position of a visible character, or the position right after one.
-- A range that STARTS inside a field code returns the hidden code as if it were text.
on vis(d, a, b)
	if b ≤ a then return 0
	tell application "Microsoft Word"
		set r to create range d start a end b
		set s to content of r
	end tell
	if s is missing value then return 0
	return (count of s)
end vis

on readRange(d, a, b)
	if b ≤ a then return ""
	tell application "Microsoft Word"
		set s to content of (create range d start a end b)
	end tell
	if s is missing value then return ""
	return s
end readRange

-- Position of the visible character with index k (counting from 0) in the whole document:
-- the largest p such that [0,p) has at most k visible characters.
-- Every probe is measured from the start of the document, which is always a clean position.
on posForVisible(d, k, docEnd, excess)
	if k < 0 then set k to 0
	set lo to k
	if lo > docEnd then set lo to docEnd
	if lo > 0 and (my vis(d, 0, lo)) > k then set lo to 0
	set hi to k + excess + 1
	if hi > docEnd then set hi to docEnd
	if hi < lo then set hi to lo
	-- The bound assumes one position holds at most one visible character; if that fails, search to the end
	if hi < docEnd and (my vis(d, 0, hi)) ≤ k then set hi to docEnd
	repeat while lo < hi
		set mid to (lo + hi + 1) div 2
		if (my vis(d, 0, mid)) ≤ k then
			set lo to mid
		else
			set hi to mid - 1
		end if
	end repeat
	return lo
end posForVisible

-- The smallest q from p on such that [p,q) has at least `need` visible characters,
-- that is, the position right after the need-th visible character. p must be a clean position.
on endForVisible(d, p, need, docEnd)
	if need ≤ 0 then return p
	set lo to p
	set hi to docEnd
	-- Grow exponentially first, so that not every probe reads up to the end of the document
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

-- Position of the visible character with index n counted from the clean position p:
-- the largest q such that [p,q) has at most n visible characters (hidden material before that character is skipped).
on startForVisible(d, p, n, docEnd)
	set lo to my endForVisible(d, p, n, docEnd)
	set hi to my endForVisible(d, p, n + 1, docEnd)
	repeat while lo < hi
		set mid to (lo + hi + 1) div 2
		if (my vis(d, p, mid)) ≤ n then
			set lo to mid
		else
			set hi to mid - 1
		end if
	end repeat
	return lo
end startForVisible

-- Number of fields with a begin or end mark inside [a,b)
on countFields(d, a, b)
	if b ≤ a then return 0
	tell application "Microsoft Word"
		return count of fields of (create range d start a end b)
	end tell
end countFields

-- Number of fields that enclose position x (begin mark before x, end mark at or after x).
-- Text inserted at such a position would land inside the field.
on fieldDepth(d, x, docEnd, nFields)
	if nFields is 0 then return 0
	return (my countFields(d, 0, x)) + (my countFields(d, x, docEnd)) - nFields
end fieldDepth

-- The begin mark of the field whose result starts exactly at x, or -1 when x is not the start of a field result
on fieldMarkBefore(d, x)
	set stride to 64
	repeat
		set a to x - stride
		if a < 0 then set a to 0
		set nf to my countFields(d, a, x)
		if nf > 0 then
			tell application "Microsoft Word"
				set f to field nf of (create range d start a end x)
				set markPos to (start of content of (field code of f)) - 1
				set resStart to start of content of (result range of f)
			end tell
			if resStart is x then return markPos
			return -1
		end if
		if a is 0 then return -1
		set stride to stride * 2
	end repeat
end fieldMarkBefore

-- The top-level fields with a mark inside [a,b): for each, the visible text from pS up to its result, and the result text
on fieldDetails(d, pS, a, b)
	set outRows to {}
	set nf to my countFields(d, a, b)
	set lastEnd to -1
	repeat with j from 1 to nf
		tell application "Microsoft Word"
			set f to field j of (create range d start a end b)
			set resStart to start of content of (result range of f)
			set resEnd to end of content of (result range of f)
		end tell
		if resStart ≥ lastEnd then
			set end of outRows to (my readRange(d, pS, resStart)) & (character id 29) & (my readRange(d, resStart, resEnd))
			set lastEnd to resEnd
		end if
	end repeat
	return my joinList(outRows, my recSep)
end fieldDetails

-- Number of equations that [a,b) touches
on countMath(d, a, b)
	if b ≤ a then return 0
	tell application "Microsoft Word"
		return count of math objects of (create range d start a end b)
	end tell
end countMath

-- How position x sits relative to the equations: "in" (inside one), "end" (right behind the last character of one),
-- "start" (right in front of one) or "" (clear of them). Text written at "in" or "end" becomes part of that equation.
on mathSide(d, x, docEnd)
	set mathLeft to false
	set mathRight to false
	if x > 0 then set mathLeft to ((my countMath(d, x - 1, x)) > 0)
	if x < docEnd then set mathRight to ((my countMath(d, x, x + 1)) > 0)
	if mathLeft and mathRight then return "in"
	if mathLeft then return "end"
	if mathRight then return "start"
	return ""
end mathSide

-- The range {start, end} of the equation that holds position x, or {-1, -1}
on equationAt(d, x)
	tell application "Microsoft Word"
		set nMath to count of math objects of d
	end tell
	repeat with i from 1 to nMath
		tell application "Microsoft Word"
			set eqRange to text range of (math object i of d)
			set a to start of content of eqRange
			set b to end of content of eqRange
		end tell
		if a ≤ x and x < b then return {a, b}
	end repeat
	return {-1, -1}
end equationAt

-- Wrap the text of every equation in t between ASCII 17 and ASCII 18, so that the caller can tell it from normal text.
-- Returns {number of equations marked, the text}; when anything does not add up, the text comes back unchanged.
-- Each request to Word takes several milliseconds and an equation needs four, so the caller limits how many equations it asks for.
on markEquations(d, t, nMath, docEnd)
	set tLen to count of t
	set spanStarts to {}
	set spanLens to {}
	set nBefore to 0
	set prevA to 0
	repeat with i from 1 to nMath
		tell application "Microsoft Word"
			set eqRange to text range of (math object i of d)
			set a to start of content of eqRange
			set b to end of content of eqRange
		end tell
		if a < prevA then return {0, t}
		-- Counted from the start of the previous equation, a position that is never inside a field code
		set nBefore to nBefore + (my vis(d, prevA, a))
		set eqText to my readRange(d, a, b)
		set nInside to count of eqText
		if nInside > 0 then
			if (nBefore + nInside) > tLen then return {0, t}
			if not (my textsEqual(text (nBefore + 1) thru (nBefore + nInside) of t, eqText)) then return {0, t}
			set end of spanStarts to nBefore
			set end of spanLens to nInside
		end if
		set prevA to a
	end repeat
	if (nBefore + (my vis(d, prevA, docEnd))) is not tLen then return {0, t}
	set outText to t
	repeat with j from (count of spanStarts) to 1 by -1
		set nBefore to item j of spanStarts
		set nInside to item j of spanLens
		set headText to ""
		if nBefore > 0 then set headText to text 1 thru nBefore of outText
		set midText to text (nBefore + 1) thru (nBefore + nInside) of outText
		set tailText to ""
		if (nBefore + nInside) < (count of outText) then set tailText to text (nBefore + nInside + 1) thru -1 of outText
		set outText to headText & (character id 17) & midText & (character id 18) & tailText
	end repeat
	return {(count of spanStarts), outText}
end markEquations

-- The error to report when the text read back at the located place is not the anchor text:
-- inside an equation the cause is the equation's structure, which these tools do not edit.
on anchorMismatch(d, a, b, got)
	set inMath to false
	try
		set inMath to ((my countMath(d, a, b)) > 0)
	end try
	if inMath then return "ERR" & my fieldSep & "in_equation"
	return "ERR" & my fieldSep & "verify_failed" & my fieldSep & got
end anchorMismatch

-- True when s holds any of the characters of charSet
on hasAnyOf(s, charSet)
	considering case, diacriticals, hyphens, punctuation and white space
		repeat with ch in (characters of charSet)
			if s contains (contents of ch) then return true
		end repeat
	end considering
	return false
end hasAnyOf

-- Where to insert between two visible characters of the located text [pS,pE): prefixText is the part of findText in front of the point.
-- Checks the text on both sides, then picks a position outside every field.
-- Returns {status, position, detail}; status is "ok", "verify_failed" or "in_field".
on insertionPoint(d, pS, pE, findText, prefixText, docEnd, nFields, allowFields)
	set segOff to count of prefixText
	if segOff > 0 then
		set pA to my endForVisible(d, pS, segOff, docEnd)
	else
		set pA to pS
	end if
	set pB to my startForVisible(d, pS, segOff, docEnd)
	set restText to ""
	if segOff < (count of findText) then set restText to text (segOff + 1) thru -1 of findText
	set gotA to my readRange(d, pS, pA)
	if not (my textsEqual(gotA, prefixText)) then return {"verify_failed", -1, gotA}
	set gotB to my readRange(d, pB, pE)
	if not (my textsEqual(gotB, restText)) then return {"verify_failed", -1, gotB}
	set p to -1
	if nFields is 0 then
		set p to pA
	else if segOff > 0 then
		-- After the preceding text: right behind it, else just past a field's end mark, else in front of the next visible character
		set cands to {pA}
		if pA < pB then set cands to {pA, pA + 1, pB}
		repeat with cand in cands
			if (my fieldDepth(d, (contents of cand), docEnd, nFields)) is 0 then
				set p to (contents of cand)
				exit repeat
			end if
		end repeat
	else
		-- In front of the anchor: at its first character, or in front of the field when the anchor starts with a field's result
		if (my fieldDepth(d, pS, docEnd, nFields)) is 0 then
			set p to pS
		else
			set markPos to my fieldMarkBefore(d, pS)
			if markPos ≥ 0 then
				if (my fieldDepth(d, markPos, docEnd, nFields)) is 0 then set p to markPos
			end if
		end if
	end if
	if p < 0 then
		if not allowFields then return {"in_field", -1, ""}
		set p to pA
	end if
	return {"ok", p, ""}
end insertionPoint

-- Locate the occ-th occurrence of findText in the document (occ=0 means it must be unique).
-- Returns {status, number of occurrences, visible offset k, full text t, docEnd, excess}
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
				-- Every open document with its location (a POSIX path, or a URL for a cloud document), then the same per window.
				-- Two open documents with one name both resolve to the same document by name or index; only their windows tell them apart.
				set docRows to {}
				set winRows to {}
				tell application "Microsoft Word"
					if not running then return "OK" & my fieldSep
					set nDocs to count of documents
					repeat with i from 1 to nDocs
						set docPlace to ""
						try
							set docPlace to posix full name of document i
						end try
						-- After one of two same-named documents is closed, Word no longer finds the other one by name
						set reachable to "1"
						try
							if not (exists document (name of document i)) then set reachable to "0"
						on error
							set reachable to "0"
						end try
						set end of docRows to (name of document i) & my fieldSep & docPlace & my fieldSep & reachable
					end repeat
					repeat with i from 1 to (count of windows)
						try
							-- Ask through the window in one specifier: a document reference held in a variable resolves by name again
							set end of winRows to (name of document of window i) & my fieldSep & (posix full name of document of window i)
						end try
					end repeat
				end tell
				return "OK" & my fieldSep & my joinList(docRows, my recSep) & (character id 29) & my joinList(winRows, my recSep)
			end if

			-- argv item 2 is the document's name, optionally followed by ASCII 29 and the location it must have
			set AppleScript's text item delimiters to (character id 29)
			set specParts to text items of (item 2 of argv)
			set AppleScript's text item delimiters to ""
			set docName to item 1 of specParts
			set wantPlace to ""
			if (count of specParts) > 1 then set wantPlace to item 2 of specParts
			tell application "Microsoft Word"
				if not (exists document docName) then
					-- Not open, or open but no longer reachable by name
					repeat with i from 1 to (count of documents)
						try
							if (name of document i) is docName then return "ERR" & my fieldSep & "lost_document"
						end try
					end repeat
					return "ERR" & my fieldSep & "not_open"
				end if
				set openNames to name of every document
			end tell
			-- Word cannot address two open documents of the same name separately, so refuse both
			set sameName to 0
			repeat with nm in openNames
				if (contents of nm) is docName then set sameName to sameName + 1
			end repeat
			if sameName > 1 then return "ERR" & my fieldSep & "duplicate_name" & my fieldSep & sameName
			tell application "Microsoft Word"
				set d to document docName
				set havePlace to ""
				try
					set havePlace to posix full name of d
				end try
			end tell
			if wantPlace is not "" then
				if not (my textsEqual(havePlace, wantPlace)) then return "ERR" & my fieldSep & "wrong_file"
			end if

			if op is "read" then
				tell application "Microsoft Word"
					set t to content of text object of d
					set tr to track revisions of d
					set nr to count of revisions of d
					set nc to count of Word comments of d
					set sv to saved of d
					set nMath to count of math objects of d
					set docEnd to end of content of text object of d
				end tell
				-- argv: read doc [markLimit]. An equation reads as its characters with a paragraph mark at every structure boundary.
				-- With a markLimit, and no more equations than that, each equation's text is marked for the caller.
				set nMarked to 0
				set markLimit to 0
				if (count of argv) > 2 then set markLimit to (item 3 of argv) as integer
				if nMath > 0 and nMath ≤ markLimit then set {nMarked, t} to my markEquations(d, t, nMath, docEnd)
				return "OK" & my fieldSep & (tr as text) & my fieldSep & nr & my fieldSep & nc & my fieldSep & (sv as text) & my fieldSep & nMath & my fieldSep & nMarked & my fieldSep & t

			else if op is "edit" then
				-- argv: edit doc findText occurrence author allowFields expectAfter nSeg  then, per segment: prefixText oldText newText
				-- A segment replaces oldText (which follows prefixText inside findText) with newText; an empty oldText is a pure insertion.
				-- Segments come in ascending order. Everything is located and checked before the first write.
				set findText to item 3 of argv
				set occ to (item 4 of argv) as integer
				set authorName to item 5 of argv
				set allowFields to ((item 6 of argv) is "1")
				set expectAfter to item 7 of argv
				set nSeg to (item 8 of argv) as integer
				set {locStatus, n, k, t, docEnd, excess} to my locateVisible(d, findText, occ)
				if locStatus is not "ok" then return "ERR" & my fieldSep & locStatus & my fieldSep & n
				set findLen to count of findText
				set pS to my posForVisible(d, k, docEnd, excess)
				set pE to my endForVisible(d, pS, findLen, docEnd)
				set got to my readRange(d, pS, pE)
				if not (my textsEqual(got, findText)) then return my anchorMismatch(d, pS, pE, got)
				set nFields to my countFields(d, 0, docEnd)
				tell application "Microsoft Word"
					set nMath to count of math objects of d
				end tell
				set segStarts to {}
				set segEnds to {}
				set segTexts to {}
				set segNearMath to {}
				set lastQ to -1
				repeat with i from 1 to nSeg
					set prefixText to item (9 + (i - 1) * 3) of argv
					set oldText to item (10 + (i - 1) * 3) of argv
					set newText to item (11 + (i - 1) * 3) of argv
					set segOff to count of prefixText
					set oldLen to count of oldText
					if oldLen > 0 then
						-- Replacement: from the first visible character of oldText to right after its last one
						set p to my startForVisible(d, pS, segOff, docEnd)
						set q to my endForVisible(d, p, oldLen, docEnd)
						set got to my readRange(d, p, q)
						if not (my textsEqual(got, oldText)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & got
						if nFields > 0 then
							set dP to my fieldDepth(d, p, docEnd, nFields)
							set dQ to my fieldDepth(d, q, docEnd, nFields)
							set p2 to p
							set q2 to q
							set dP2 to dP
							set dQ2 to dQ
							-- oldText ends with a field's result: q stops at the field's end mark, so step over the mark
							repeat 3 times
								if dQ2 is 0 or q2 ≥ docEnd then exit repeat
								if (my vis(d, q2, q2 + 1)) > 0 then exit repeat
								set q2 to q2 + 1
								set dQ2 to my fieldDepth(d, q2, docEnd, nFields)
							end repeat
							-- oldText starts with a field's result: start at the field's begin mark instead
							if dP2 > 0 then
								set markPos to my fieldMarkBefore(d, p)
								if markPos ≥ 0 and markPos ≥ lastQ then
									set p2 to markPos
									set dP2 to my fieldDepth(d, p2, docEnd, nFields)
								end if
							end if
							if dP2 is 0 and dQ2 is 0 then
								if (my countFields(d, p2, q2)) > 0 then
									-- Whole fields lie inside the text to be replaced
									if not allowFields then return "ERR" & my fieldSep & "field_in_range" & my fieldSep & (my fieldDetails(d, pS, p2, q2))
									set p to p2
									set q to q2
								end if
							else
								-- The text is inside a field's result, or cuts through a field
								if not (allowFields and dP is dQ and (my countFields(d, p, q)) is 0) then return "ERR" & my fieldSep & "in_field" & my fieldSep & got
							end if
						end if
						set nearMath to false
						if nMath > 0 then
							-- Rewriting text that belongs to an equation would turn the equation into plain characters
							if (my countMath(d, p, q)) > 0 then return "ERR" & my fieldSep & "equation_in_range" & my fieldSep & got
							set nearMath to ((my mathSide(d, p, docEnd)) is not "")
						end if
					else
						-- Pure insertion between two visible characters
						set {insStatus, p, insDetail} to my insertionPoint(d, pS, pE, findText, prefixText, docEnd, nFields, allowFields)
						if insStatus is not "ok" then return "ERR" & my fieldSep & insStatus & my fieldSep & insDetail
						set q to p
						set nearMath to false
						if nMath > 0 then
							set mSide to my mathSide(d, p, docEnd)
							if mSide is "in" then return "ERR" & my fieldSep & "in_equation"
							if mSide is "end" then
								-- Text written at the end of an equation joins it. Rewrite the next character instead: what replaces normal text stays normal text.
								set nextLen to 1
								set nextChar to my readRange(d, p, p + 1)
								if nextChar is not "" then
									if (id of nextChar) ≥ 55296 and (id of nextChar) ≤ 56319 then
										set nextLen to 2
										set nextChar to my readRange(d, p, p + 2)
									end if
								end if
								if nextChar is "" then return "ERR" & my fieldSep & "after_equation"
								if (count of nextChar) is not 1 then return "ERR" & my fieldSep & "after_equation"
								if (id of nextChar) < 32 then return "ERR" & my fieldSep & "after_equation"
								if (my countMath(d, p, p + nextLen)) > 0 then return "ERR" & my fieldSep & "in_equation"
								if nFields > 0 then
									if (my fieldDepth(d, p, docEnd, nFields)) > 0 or (my fieldDepth(d, p + nextLen, docEnd, nFields)) > 0 then return "ERR" & my fieldSep & "after_equation"
								end if
								set q to p + nextLen
								set newText to newText & nextChar
							end if
							set nearMath to (mSide is not "")
						end if
					end if
					if p < lastQ then return "ERR" & my fieldSep & "bad_segments" & my fieldSep & i
					set lastQ to q
					set end of segStarts to p
					set end of segEnds to q
					set end of segTexts to newText
					set end of segNearMath to nearMath
				end repeat
				-- Write from the last segment to the first, so earlier positions stay valid
				tell application "Microsoft Word"
					set oldName to user name
					set track revisions of d to true
					set nDone to 0
					set mathLeak to false
					try
						if authorName is not "" then set user name to authorName
						repeat with i from nSeg to 1 by -1
							set r to create range d start (item i of segStarts) end (item i of segEnds)
							set content of r to (item i of segTexts)
							set nDone to nDone + 1
							-- Next to an equation: what was just written must not have become part of it
							if (item i of segNearMath) and (item i of segTexts) is not "" then
								if (count of math objects of (create range d start (item i of segStarts) end ((item i of segStarts) + 1))) > 0 then set mathLeak to true
							end if
						end repeat
						set user name to oldName
					on error errMsg
						set user name to oldName
						return "ERR" & my fieldSep & "write_failed" & my fieldSep & nDone & my fieldSep & errMsg
					end try
					set nr to count of revisions of d
					set t2 to content of text object of d
				end tell
				-- Read the whole document back: its visible text must now differ from before in exactly this one place
				set headText to ""
				if k > 0 then set headText to text 1 thru k of t
				set tailText to ""
				if (k + findLen) < (count of t) then set tailText to text (k + findLen + 1) thru -1 of t
				set postStatus to "ok"
				if not (my textsEqual(t2, headText & expectAfter & tailText)) then set postStatus to "mismatch"
				if mathLeak then set postStatus to "mismatch"
				set ctxA to ""
				if k > 0 then
					set ctxFrom to k - 59
					if ctxFrom < 1 then set ctxFrom to 1
					set ctxA to text ctxFrom thru k of t
				end if
				set ctxB to ""
				if (count of tailText) > 60 then
					set ctxB to text 1 thru 60 of tailText
				else
					set ctxB to tailText
				end if
				set posList to {}
				repeat with i from 1 to nSeg
					set end of posList to ((item i of segStarts) as text) & "-" & ((item i of segEnds) as text)
				end repeat
				return "OK" & my fieldSep & n & my fieldSep & (my joinList(posList, ",")) & my fieldSep & nr & my fieldSep & postStatus & my fieldSep & ctxA & my fieldSep & ctxB

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
				if got is missing value then set got to ""
				if not (my textsEqual(got, findText)) then return my anchorMismatch(d, p, q, got)
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
						set rvText to content of text object of rv
						if rvText is missing value then set rvText to ""
						set end of revLines to ((revision type of rv) as text) & my fieldSep & (author of rv) & my fieldSep & rvText
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

			else if op is "locate" then
				-- argv: locate doc findText occurrence  → position of the anchor text in the document
				set findText to item 3 of argv
				set occ to (item 4 of argv) as integer
				set {locStatus, n, k, t, docEnd, excess} to my locateVisible(d, findText, occ)
				if locStatus is not "ok" then return "ERR" & my fieldSep & locStatus & my fieldSep & n
				set p to my posForVisible(d, k, docEnd, excess)
				set q to my endForVisible(d, p, (count of findText), docEnd)
				set got to my readRange(d, p, q)
				if not (my textsEqual(got, findText)) then return my anchorMismatch(d, p, q, got)
				return "OK" & my fieldSep & n & my fieldSep & p & my fieldSep & q

			else if op is "images" then
				-- List the inline pictures: index, type, width and height (points), position, whether tracked-deleted, text around.
				-- The text around is cut from the document's visible text: a range that starts just before a picture may start inside a field code.
				set imgRows to {}
				tell application "Microsoft Word"
					set nShapes to count of inline shapes of d
					set nFloating to count of shapes of d
					set t to content of text object of d
				end tell
				set tLen to count of t
				repeat with i from 1 to nShapes
					tell application "Microsoft Word"
						set shp to inline shape i of d
						set shpKind to (inline shape type of shp) as text
						set shpW to width of shp
						set shpH to height of shp
						set shpStart to start of content of text object of shp
						set selfText to content of text object of shp
					end tell
					if selfText is missing value then set selfText to ""
					set selfLen to count of selfText
					set nBefore to my vis(d, 0, shpStart)
					set textAbove to ""
					if nBefore > 0 then
						set fromChar to nBefore - 149
						if fromChar < 1 then set fromChar to 1
						set textAbove to text fromChar thru nBefore of t
					end if
					set textBelow to ""
					set belowFrom to nBefore + selfLen + 1
					if belowFrom ≤ tLen then
						set belowTo to belowFrom + 259
						if belowTo > tLen then set belowTo to tLen
						set textBelow to text belowFrom thru belowTo of t
					end if
					set end of imgRows to (i as text) & my fieldSep & shpKind & my fieldSep & (shpW as text) & my fieldSep & (shpH as text) & my fieldSep & (shpStart as text) & my fieldSep & (selfLen as text) & my fieldSep & textAbove & my fieldSep & textBelow
				end repeat
				return "OK" & my fieldSep & nShapes & my fieldSep & nFloating & my recSep & my joinList(imgRows, my recSep)

			else if op is "insertimage" then
				-- argv: insertimage doc anchorText occurrence where(before|after) imagePath author widthPoints ownParagraph(1|0)
				-- Insert a new inline picture before or after the anchor text, as a tracked change.
				-- With ownParagraph the picture gets a paragraph to itself: paragraph marks are added only where none is there already.
				set findText to item 3 of argv
				set occ to (item 4 of argv) as integer
				set placeWhere to item 5 of argv
				set imgPath to item 6 of argv
				set authorName to item 7 of argv
				set wantWidth to (item 8 of argv) as real
				set ownPara to ((item 9 of argv) is "1")
				set {locStatus, n, k, t, docEnd, excess} to my locateVisible(d, findText, occ)
				if locStatus is not "ok" then return "ERR" & my fieldSep & locStatus & my fieldSep & n
				set findLen to count of findText
				set pS to my posForVisible(d, k, docEnd, excess)
				set pE to my endForVisible(d, pS, findLen, docEnd)
				set got to my readRange(d, pS, pE)
				if not (my textsEqual(got, findText)) then return my anchorMismatch(d, pS, pE, got)
				set nFields to my countFields(d, 0, docEnd)
				set prefixText to ""
				if placeWhere is "after" then set prefixText to findText
				set {insStatus, p, insDetail} to my insertionPoint(d, pS, pE, findText, prefixText, docEnd, nFields, false)
				if insStatus is not "ok" then return "ERR" & my fieldSep & insStatus & my fieldSep & insDetail
				set mSide to my mathSide(d, p, docEnd)
				if mSide is "in" then return "ERR" & my fieldSep & "in_equation"
				if mSide is "end" then return "ERR" & my fieldSep & "after_equation"
				-- kIns visible characters come before the insertion point
				set kIns to k + (count of prefixText)
				set tLen to count of t
				set leadMark to ""
				set trailMark to ""
				if ownPara then
					if kIns > 0 then
						if (id of (character kIns of t)) is not 13 then set leadMark to (character id 13)
					end if
					if kIns < tLen then
						if (id of (character (kIns + 1) of t)) is not 13 then set trailMark to (character id 13)
					end if
				end if
				set picPos to p + (count of leadMark)
				tell application "Microsoft Word"
					set textWidth to 0
					try
						set ps to page setup of d
						set textWidth to (page width of ps) - (left margin of ps) - (right margin of ps)
					end try
				end tell
				set stage to "start"
				tell application "Microsoft Word"
					set oldName to user name
					set track revisions of d to true
					try
						if authorName is not "" then set user name to authorName
						if (leadMark & trailMark) is not "" then
							set content of (create range d start p end p) to (leadMark & trailMark)
							set stage to "marks"
						end if
						make new inline picture at (create range d start picPos end picPos) with properties {file name:imgPath, link to file:false, save with document:true}
						set stage to "inserted"
						-- Find the new picture by its position
						set shpIndex to 0
						set nShapes to count of inline shapes of d
						repeat with i from 1 to nShapes
							if (start of content of text object of inline shape i of d) is picPos then
								set shpIndex to i
								exit repeat
							end if
						end repeat
						if shpIndex is 0 then error "the new picture was not found at the expected position after inserting it"
						set newShp to inline shape shpIndex of d
						set natW to width of newShp
						set natH to height of newShp
						set finalW to natW
						if wantWidth > 0 then
							set finalW to wantWidth
						else if textWidth > 0 and natW > textWidth then
							set finalW to textWidth
						end if
						if finalW is not natW then
							try
								set lock aspect ratio of newShp to true
							end try
							set width of newShp to finalW
							set height of newShp to (finalW * natH / natW)
						end if
						set stage to "sized"
						if ownPara then
							-- A paragraph with exact line spacing would crop the picture to one line
							try
								set pf to paragraph format of (create range d start picPos end (picPos + 1))
								if (line spacing rule of pf) is line space exactly then set line spacing rule of pf to line space single
							end try
						end if
						set user name to oldName
					on error errMsg
						set user name to oldName
						return "ERR" & my fieldSep & "insert_image_failed" & my fieldSep & stage & my fieldSep & errMsg
					end try
					set finalShp to inline shape shpIndex of d
					set newW to width of finalShp
					set newH to height of finalShp
					set nRev to count of revisions of d
					set t2 to content of text object of d
				end tell
				-- Read the document back: its visible text must now differ only by the picture and the paragraph marks added for it
				set picText to my readRange(d, picPos, picPos + 1)
				set headText to ""
				if kIns > 0 then set headText to text 1 thru kIns of t
				set tailText to ""
				if kIns < tLen then set tailText to text (kIns + 1) thru -1 of t
				set postStatus to "ok"
				if picText is "" or not (my textsEqual(t2, headText & leadMark & picText & trailMark & tailText)) then set postStatus to "mismatch"
				set ctxA to ""
				if kIns > 0 then
					set ctxFrom to kIns - 59
					if ctxFrom < 1 then set ctxFrom to 1
					set ctxA to text ctxFrom thru kIns of t
				end if
				set ctxB to tailText
				if (count of tailText) > 60 then set ctxB to text 1 thru 60 of tailText
				return "OK" & my fieldSep & shpIndex & my fieldSep & (natW as text) & my fieldSep & (natH as text) & my fieldSep & (newW as text) & my fieldSep & (newH as text) & my fieldSep & nRev & my fieldSep & postStatus & my fieldSep & (count of leadMark) & my fieldSep & (count of trailMark) & my fieldSep & ctxA & my fieldSep & ctxB

			else if op is "insertequation" then
				-- argv: insertequation doc anchorText occurrence where(before|after) linearText author ownParagraph(1|0) unbuiltChars
				-- Write the formula in Word's linear format before or after the anchor text, turn that text into an equation and build it up; all tracked.
				-- unbuiltChars only remain in an equation's text when Word could not build it; the insertion is then taken out again.
				set findText to item 3 of argv
				set occ to (item 4 of argv) as integer
				set placeWhere to item 5 of argv
				set eqText to item 6 of argv
				set authorName to item 7 of argv
				set ownPara to ((item 8 of argv) is "1")
				set unbuiltChars to item 9 of argv
				set {locStatus, n, k, t, docEnd, excess} to my locateVisible(d, findText, occ)
				if locStatus is not "ok" then return "ERR" & my fieldSep & locStatus & my fieldSep & n
				set findLen to count of findText
				set pS to my posForVisible(d, k, docEnd, excess)
				set pE to my endForVisible(d, pS, findLen, docEnd)
				set got to my readRange(d, pS, pE)
				if not (my textsEqual(got, findText)) then return my anchorMismatch(d, pS, pE, got)
				set nFields to my countFields(d, 0, docEnd)
				set prefixText to ""
				if placeWhere is "after" then set prefixText to findText
				set {insStatus, p, insDetail} to my insertionPoint(d, pS, pE, findText, prefixText, docEnd, nFields, false)
				if insStatus is not "ok" then return "ERR" & my fieldSep & insStatus & my fieldSep & insDetail
				-- An equation written right next to another one would merge with it
				set mSide to my mathSide(d, p, docEnd)
				if mSide is not "" then return "ERR" & my fieldSep & "near_equation" & my fieldSep & mSide
				-- kIns visible characters come before the insertion point
				set kIns to k + (count of prefixText)
				set tLen to count of t
				set leadMark to ""
				set trailMark to ""
				if ownPara then
					if kIns > 0 then
						if (id of (character kIns of t)) is not 13 then set leadMark to (character id 13)
					end if
					if kIns < tLen then
						if (id of (character (kIns + 1) of t)) is not 13 then set trailMark to (character id 13)
					end if
				end if
				set nLead to count of leadMark
				set nTrail to count of trailMark
				set stage to "start"
				set errText to ""
				set eqA to p + nLead
				tell application "Microsoft Word"
					set oldName to user name
					set track revisions of d to true
					set nMath0 to count of math objects of d
					set e0 to end of content of text object of d
					try
						if authorName is not "" then set user name to authorName
						set content of (create range d start p end p) to (leadMark & eqText & trailMark)
						set e1 to end of content of text object of d
						set eqB to p + (e1 - e0) - nTrail
						set stage to "text"
						create new equation from range (create range d start eqA end eqB) in document d
						set stage to "equation"
						set e2 to end of content of text object of d
						build up (math object 1 of (create range d start eqA end (eqB + (e2 - e1))))
						set stage to "built"
					on error errMsg
						set errText to errMsg
					end try
					set user name to oldName
					set e3 to end of content of text object of d
				end tell
				-- Everything written now spans [p, p + grown); the equation itself is [eqA, zB)
				set grown to e3 - e0
				set zB to p + grown - nTrail
				set eqRead to ""
				if stage is not "start" then set eqRead to my readRange(d, eqA, zB)
				set unbuilt to false
				if stage is "built" then set unbuilt to my hasAnyOf(eqRead, unbuiltChars)
				if unbuilt or stage is not "built" then
					-- Take the insertion out again: deleting one's own pending insertion removes it without a trace
					set cleaned to "1"
					if grown > 0 then
						tell application "Microsoft Word"
							try
								if authorName is not "" then set user name to authorName
								set content of (create range d start p end (p + grown)) to ""
							end try
							set user name to oldName
							set t3 to content of text object of d
						end tell
						if not (my textsEqual(t3, t)) then set cleaned to "0"
					end if
					if unbuilt then return "ERR" & my fieldSep & "equation_unbuilt" & my fieldSep & cleaned & my fieldSep & eqRead
					return "ERR" & my fieldSep & "equation_failed" & my fieldSep & stage & my fieldSep & cleaned & my fieldSep & errText
				end if
				tell application "Microsoft Word"
					set nMath1 to count of math objects of d
					set nRev to count of revisions of d
					set t2 to content of text object of d
				end tell
				-- Read the document back: its visible text must now differ only by the equation and the paragraph marks added for it
				set headText to ""
				if kIns > 0 then set headText to text 1 thru kIns of t
				set tailText to ""
				if kIns < tLen then set tailText to text (kIns + 1) thru -1 of t
				set postStatus to "ok"
				if eqRead is "" or not (my textsEqual(t2, headText & leadMark & eqRead & trailMark & tailText)) then set postStatus to "mismatch"
				if postStatus is "ok" and nMath1 is not (nMath0 + 1) then set postStatus to "merged"
				set ctxA to ""
				if kIns > 0 then
					set ctxFrom to kIns - 59
					if ctxFrom < 1 then set ctxFrom to 1
					set ctxA to text ctxFrom thru kIns of t
				end if
				set ctxB to tailText
				if (count of tailText) > 60 then set ctxB to text 1 thru 60 of tailText
				return "OK" & my fieldSep & nMath1 & my fieldSep & nRev & my fieldSep & postStatus & my fieldSep & nLead & my fieldSep & nTrail & my fieldSep & eqRead & my fieldSep & ctxA & my fieldSep & ctxB

			else if op is "deleteequation" then
				-- argv: deleteequation doc anchorText occurrence where(before|after) author
				-- Delete, as a tracked change, the equation that directly follows (after) or precedes (before) the anchor text.
				-- Paragraph marks and spaces between the two are skipped. An equation that fills a paragraph goes together with its paragraph mark.
				set findText to item 3 of argv
				set occ to (item 4 of argv) as integer
				set placeWhere to item 5 of argv
				set authorName to item 6 of argv
				set {locStatus, n, k, t, docEnd, excess} to my locateVisible(d, findText, occ)
				if locStatus is not "ok" then return "ERR" & my fieldSep & locStatus & my fieldSep & n
				set findLen to count of findText
				set pS to my posForVisible(d, k, docEnd, excess)
				set pE to my endForVisible(d, pS, findLen, docEnd)
				set got to my readRange(d, pS, pE)
				if not (my textsEqual(got, findText)) then return my anchorMismatch(d, pS, pE, got)
				set tLen to count of t
				-- kx: index (from 0) of the visible character next to the anchor on that side
				if placeWhere is "after" then
					set kx to k + findLen
					set stepBy to 1
				else
					set kx to k - 1
					set stepBy to -1
				end if
				set hitPos to -1
				repeat 4 times
					if kx < 0 or kx ≥ tLen then exit repeat
					set x to my posForVisible(d, kx, docEnd, excess)
					if (my countMath(d, x, x + 1)) > 0 then
						set hitPos to x
						exit repeat
					end if
					set chId to id of (character (kx + 1) of t)
					if not (chId is 13 or chId is 32 or chId is 160) then exit repeat
					set kx to kx + stepBy
				end repeat
				if hitPos < 0 then return "ERR" & my fieldSep & "no_equation"
				set {zA, zB} to my equationAt(d, hitPos)
				if zA < 0 then return "ERR" & my fieldSep & "no_equation"
				set eqRead to my readRange(d, zA, zB)
				if eqRead is "" then return "ERR" & my fieldSep & "no_equation"
				-- The equation's text must sit at that place in the document's text
				set kA to my vis(d, 0, zA)
				set eqLen to count of eqRead
				if (kA + eqLen) > tLen then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & eqRead
				if not (my textsEqual(text (kA + 1) thru (kA + eqLen) of t, eqRead)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & eqRead
				-- Alone in its paragraph: take the paragraph mark behind it along, so that no empty paragraph is left.
				-- (The mark in front stays: removing that one would give the paragraph above the formatting of the equation's paragraph.)
				set dropMark to 0
				set aloneBefore to (kA is 0)
				if kA > 0 then set aloneBefore to ((id of (character kA of t)) is 13)
				if aloneBefore and (kA + eqLen) < tLen and (zB + 1) < docEnd then
					if (id of (character (kA + eqLen + 1) of t)) is 13 then
						if (my readRange(d, zB, zB + 1)) is (character id 13) then set dropMark to 1
					end if
				end if
				set errText to ""
				tell application "Microsoft Word"
					set oldName to user name
					set track revisions of d to true
					set nMath0 to count of math objects of d
					try
						if authorName is not "" then set user name to authorName
						set content of (create range d start zA end (zB + dropMark)) to ""
					on error errMsg
						set errText to errMsg
					end try
					set user name to oldName
					set nMath1 to count of math objects of d
					set nRev to count of revisions of d
					set t2 to content of text object of d
				end tell
				if errText is not "" then return "ERR" & my fieldSep & "write_failed" & my fieldSep & "0" & my fieldSep & errText
				set headText to ""
				if kA > 0 then set headText to text 1 thru kA of t
				set tailText to ""
				if (kA + eqLen + dropMark) < tLen then set tailText to text (kA + eqLen + dropMark + 1) thru -1 of t
				set postStatus to "ok"
				if not (my textsEqual(t2, headText & tailText)) then set postStatus to "mismatch"
				set ctxA to headText
				if (count of headText) > 60 then set ctxA to text -60 thru -1 of headText
				set ctxB to tailText
				if (count of tailText) > 60 then set ctxB to text 1 thru 60 of tailText
				return "OK" & my fieldSep & nMath0 & my fieldSep & nMath1 & my fieldSep & nRev & my fieldSep & postStatus & my fieldSep & dropMark & my fieldSep & eqRead & my fieldSep & ctxA & my fieldSep & ctxB

			else if op is "replaceimage" then
				-- argv: replaceimage doc index expectedStart imagePath author
				-- Insert the new picture at the same position (old width, height from the new proportions), then delete the old one; both steps are tracked.
				set shpIndex to (item 3 of argv) as integer
				set expectedStart to (item 4 of argv) as integer
				set imgPath to item 5 of argv
				set authorName to item 6 of argv
				tell application "Microsoft Word"
					set nShapes to count of inline shapes of d
				end tell
				if shpIndex < 1 or shpIndex > nShapes then return "ERR" & my fieldSep & "badindex" & my fieldSep & nShapes
				tell application "Microsoft Word"
					set oldShp to inline shape shpIndex of d
					set oldStart to start of content of text object of oldShp
					set oldW to width of oldShp
					set oldH to height of oldShp
					set oldAlt to ""
					try
						set oldAlt to alternative text of oldShp
						if oldAlt is missing value then set oldAlt to ""
					end try
				end tell
				if oldStart is not expectedStart then return "ERR" & my fieldSep & "stale" & my fieldSep & oldStart
				set stage to "start"
				tell application "Microsoft Word"
					set oldName to user name
					set track revisions of d to true
					try
						if authorName is not "" then set user name to authorName
						set insRange to create range d start oldStart end oldStart
						make new inline picture at insRange with properties {file name:imgPath, link to file:false, save with document:true}
						set stage to "inserted"
						-- The new picture now has the old index; the old picture has moved up by one
						set newShp to inline shape shpIndex of d
						if (start of content of text object of newShp) is not oldStart then error "the new picture was not found at the expected position after inserting it"
						set natW to width of newShp
						set natH to height of newShp
						try
							set lock aspect ratio of newShp to true
						end try
						set width of newShp to oldW
						set height of newShp to (oldW * natH / natW)
						if oldAlt is not "" then
							try
								set alternative text of newShp to oldAlt
							end try
						end if
						set stage to "resized"
						set oldShp2 to inline shape (shpIndex + 1) of d
						set oldStart2 to start of content of text object of oldShp2
						if oldStart2 is not (oldStart + 1) then error "the old picture was not found at the expected position after inserting the new one"
						set content of (create range d start oldStart2 end (oldStart2 + 1)) to ""
						set stage to "done"
						set user name to oldName
					on error errMsg
						set user name to oldName
						return "ERR" & my fieldSep & "image_failed" & my fieldSep & stage & my fieldSep & errMsg
					end try
					set finalShp to inline shape shpIndex of d
					set newW to width of finalShp
					set newH to height of finalShp
					set nRev to count of revisions of d
				end tell
				return "OK" & my fieldSep & (oldW as text) & my fieldSep & (oldH as text) & my fieldSep & (natW as text) & my fieldSep & (natH as text) & my fieldSep & (newW as text) & my fieldSep & (newH as text) & my fieldSep & nRev
			end if
			return "ERR" & my fieldSep & "unknown_op"
		end timeout
	on error errMsg number errNum
		return "ERR" & my fieldSep & "exception" & my fieldSep & errMsg & " (" & errNum & ")"
	end try
end run
