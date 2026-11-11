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
				end tell
				return "OK" & my fieldSep & (tr as text) & my fieldSep & nr & my fieldSep & nc & my fieldSep & (sv as text) & my fieldSep & t

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
				if not (my textsEqual(got, findText)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & got
				set nFields to my countFields(d, 0, docEnd)
				set segStarts to {}
				set segEnds to {}
				set segTexts to {}
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
					else
						-- Pure insertion between two visible characters: check the text on both sides, then pick a position outside every field
						if segOff > 0 then
							set pA to my endForVisible(d, pS, segOff, docEnd)
						else
							set pA to pS
						end if
						set pB to my startForVisible(d, pS, segOff, docEnd)
						set restText to ""
						if segOff < findLen then set restText to text (segOff + 1) thru -1 of findText
						if not (my textsEqual(my readRange(d, pS, pA), prefixText)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & (my readRange(d, pS, pA))
						if not (my textsEqual(my readRange(d, pB, pE), restText)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & (my readRange(d, pB, pE))
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
							if not allowFields then return "ERR" & my fieldSep & "in_field" & my fieldSep & ""
							set p to pA
						end if
						set q to p
					end if
					if p < lastQ then return "ERR" & my fieldSep & "bad_segments" & my fieldSep & i
					set lastQ to q
					set end of segStarts to p
					set end of segEnds to q
					set end of segTexts to newText
				end repeat
				-- Write from the last segment to the first, so earlier positions stay valid
				tell application "Microsoft Word"
					set oldName to user name
					set track revisions of d to true
					set nDone to 0
					try
						if authorName is not "" then set user name to authorName
						repeat with i from nSeg to 1 by -1
							set r to create range d start (item i of segStarts) end (item i of segEnds)
							set content of r to (item i of segTexts)
							set nDone to nDone + 1
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
				if not (my textsEqual(got, findText)) then return "ERR" & my fieldSep & "verify_failed" & my fieldSep & got
				return "OK" & my fieldSep & n & my fieldSep & p & my fieldSep & q

			else if op is "images" then
				-- List the inline pictures: index, type, width and height (points), position, whether tracked-deleted, text around
				set imgRows to {}
				tell application "Microsoft Word"
					set nShapes to count of inline shapes of d
					set nFloating to count of shapes of d
					set docEnd to end of content of text object of d
				end tell
				repeat with i from 1 to nShapes
					tell application "Microsoft Word"
						set shp to inline shape i of d
						set shpKind to (inline shape type of shp) as text
						set shpW to width of shp
						set shpH to height of shp
						set shpStart to start of content of text object of shp
						set shpEnd to end of content of text object of shp
						set selfText to content of text object of shp
						set ctxA to shpStart - 150
						if ctxA < 0 then set ctxA to 0
						set ctxB to shpEnd + 260
						if ctxB > docEnd then set ctxB to docEnd
						set textAbove to ""
						if shpStart > ctxA then set textAbove to content of (create range d start ctxA end shpStart)
						set textBelow to ""
						if ctxB > shpEnd then set textBelow to content of (create range d start shpEnd end ctxB)
					end tell
					if selfText is missing value then set selfText to ""
					if textAbove is missing value then set textAbove to ""
					if textBelow is missing value then set textBelow to ""
					set end of imgRows to (i as text) & my fieldSep & shpKind & my fieldSep & (shpW as text) & my fieldSep & (shpH as text) & my fieldSep & (shpStart as text) & my fieldSep & ((count of selfText) as text) & my fieldSep & textAbove & my fieldSep & textBelow
				end repeat
				return "OK" & my fieldSep & nShapes & my fieldSep & nFloating & my recSep & my joinList(imgRows, my recSep)

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
