﻿// Copyright 2017 Mike Nordell
// The following code is 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
//using Microsoft.CodeAnalysis.Text;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;

namespace HotCommands.Commands.ProjectLess
{
	internal static class NoProj_ExpandSelection
	{
		/// <summary>
		/// <para>Performs a non-syntactical expansion of selection, to matching: '', "", (), {} and [].</para>
		/// <para>If current selection should be something like (using | to delimit selection):</para>
		/// <para>( |(xx(yy)zz| ) )</para>
		/// <para>it will result in the non-syntactically correct selection:</para>
		/// <para>|( (xx(yy)zz )| )</para>
		/// <para>TODO (in no special order):</para>
		/// <para>Multi-line</para>
		/// <para>Language-specific implementations for syntactical correctness, including but not limited to: C, C++, C#.</para>
		/// </summary>
		/// <param name="filter"></param>
		/// <param name="textView"></param>
		/// <returns></returns>
		/// <remarks>
		/// <para>The algorithm currently is as follows:</para>
		/// <para>From selection start, it walks towards the start of the line, looking for a delimiter.</para>
		/// <para>It then tries to find a matching delimiter from selection end towards the end of the line.</para>
		/// <para>If no matching delimiters have been found when reaching the first pos of the line, it gives up.</para>
		/// </remarks>
		public static int HandleCommandExpandTask(CommandFilter filter, IWpfTextView textView)
		{
			if (textView.Selection.IsEmpty)
			{
				filter.editorOperations.SelectCurrentWord(); // Let the default "Edit.SelectCurrentWord" handle it.
				return VSConstants.S_OK;
			}
			#region Precondition checks
			if (filter == null || filter.editorOperations == null) // should be impossible?
			{
				goto L_abort;
			}
			if (textView.Selection.Mode != TextSelectionMode.Stream)
			{
				goto L_abort; // We don't yet handle multi-line
			}
			var initSelPos1 = textView.Selection.Start.Position;
			var initSelPos2 = textView.Selection.End.Position;
			ITextSnapshotLine line1 = initSelPos1.GetContainingLine();
			ITextSnapshotLine line2 = initSelPos2.GetContainingLine();
			if (line1.LineNumber != line2.LineNumber)
			{
				goto L_abort; // We don't yet handle multi-line
			}
#if _DEBUG
			if (initSelPos1 <= line1.Start || initSelPos2 >= line2.End)
			{
				goto L_abort; // something is seriously wrong
			}
#endif
			#endregion // Precondition checks
			// TODO: This is some hairy stuff. There must be a better solution than this brick.
			bool pos1JustInsideDelimiter = initSelPos1 > line1.Start && IsLeadingDelimiter(initSelPos1.Subtract(1).GetChar());
			bool pos2JustInsideDelimiter = initSelPos2 < line2.End   && IsTrailingDelimiter(initSelPos2.GetChar());
			bool pos1IsOnDelimiter       = IsLeadingDelimiter(initSelPos1.GetChar());
			bool pos2IsOnDelimiter       = IsTrailingDelimiter(initSelPos2.Subtract(1).GetChar());
			bool areOnMatchingDelimiters = pos1IsOnDelimiter && pos2IsOnDelimiter && IsMatchingDelimiters(initSelPos1.GetChar(), initSelPos2.Subtract(1).GetChar());
			int startAdjust = areOnMatchingDelimiters ? 1 : 0;
			SnapshotPoint start1 = initSelPos1.Subtract(initSelPos1 > line1.Start ? startAdjust : 0);
			SnapshotPoint start2 = initSelPos2.Add     (initSelPos2 < line2.End   ? startAdjust : 0);
			for (SnapshotPoint tstPos1 = start1; tstPos1 >= line1.Start; tstPos1 = tstPos1.Subtract(1))
			{
				char ch1 = tstPos1.GetChar();
				if (!IsLeadingDelimiter(ch1)) { continue; }
				for (SnapshotPoint tstPos2 = start2; tstPos2 < line2.End; tstPos2 = tstPos2.Add(1))
				{
					char ch2 = tstPos2.GetChar();
					if (!IsMatchingDelimiters(ch1, ch2)) { continue; }
					// For Select(); pos 1 is the position of the first char, but pos2 is the position _after_ the last char.
					if ((pos1JustInsideDelimiter && pos2JustInsideDelimiter && tstPos1 != start1) ||
						(pos1IsOnDelimiter && tstPos1 == start1))
					{
						tstPos2 = tstPos2.Add(1); // include the delimiters in the selection
					}
					else
					{
						tstPos1 = tstPos1.Add(1); // select only what's inside the delimiters
					}
					textView.Selection.Select(new SnapshotSpan(tstPos1, tstPos2), false);
					return VSConstants.S_OK;
				}
				if (tstPos1 == line1.Start) { break; } // Unfortunately need this due to VS' API.
			}
		L_abort:
			return VSConstants.E_ABORT; // TODO: Is there a more appropriate error code?
		}

		static bool IsMatchingDelimiters(char ch1, char ch2)
		{
			switch (ch1)
			{
				case '\'':
				case '"': return ch2 == ch1;
				case '(': return ch2 == ')';
				case '[': return ch2 == ']';
				case '{': return ch2 == '}';
			}
			return false;
		}
		static bool IsLeadingDelimiter(char ch)
		{
			switch (ch)
			{
				case '\'':
				case '"':
				case '(':
				case '{':
				case '[':
					return true;
			}
			return false;
		}
		static bool IsTrailingDelimiter(char ch)
		{
			switch (ch)
			{
				case '\'':
				case '"':
				case ')':
				case '}':
				case ']':
					return true;
			}
			return false;
		}
	}
}
