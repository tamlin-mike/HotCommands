// Copyright 2017 Mike Nordell
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
		public static int HandleCommandExpandTask(IEditorOperations editorOps, IWpfTextView textView)
		{
			#region Precondition checks
			if (editorOps == null) // should be impossible?
			{
				goto L_abort;
			}
			if (textView.Selection.IsEmpty)
			{
				editorOps.SelectCurrentWord(); // Let the default "Edit.SelectCurrentWord" handle it.
				return VSConstants.S_OK;
			}
			if (textView.Selection.Mode != TextSelectionMode.Stream)
			{
				goto L_abort; // Expand in Box mode seems, as Spock would have said, highly illogical.
			}
			#endregion // Precondition checks
			var initSelPos1 = textView.Selection.Start.Position;
			var initSelPos2 = textView.Selection.End.Position;
			ITextSnapshot shot = initSelPos1.Snapshot;
			// Current intended design:
			// - If pos1 is just outside a delimiter, but pos2 is not, try to extend the selection by only moving pos2.   Delimiters are included in selection.
			// - If pos2 is just outside a delimiter, but pos1 is not, try to extend the selection by only moving pos1.   Delimiters are included in selection.
			// - If both pos1 and pos2 are just inside matching delimiters, extend the selection by 1 in both directions. Delimiters are included in selection.
			// - If pos1 is just inside a delimiter, but pos2 is not, try to extend the selection by only moving pos2.    Delimiters not included in selection.
			// - If pos2 is just inside a delimiter, but pos1 is not, try to extend the selection by only moving pos1.    Delimiters not included in selection.
			// - In any other case, seek pos1 towards start to find a delimiter, then a matching delimiter for pos2.      Delimiters not included in selection.
			// TODO: This is some hairy stuff. There must be a better solution than this brick.
			int iSelPos1 = initSelPos1.Position;
			int iSelPos2 = initSelPos2.Position;
			int selLen = shot.Length;
			bool canMovePos1 = iSelPos1 > 0;
			bool canMovePos2 = iSelPos2 < selLen;
			bool pos1JustInsideDelimiter = canMovePos1 && IsLeadingDelimiter(shot[iSelPos1 - 1]);
			bool pos2JustInsideDelimiter = canMovePos2 && IsTrailingDelimiter(shot[iSelPos2]);
			bool bothJustInsideDelimiter = pos1JustInsideDelimiter && pos2JustInsideDelimiter;
			bool pos1IsOnDelimiter       = IsLeadingDelimiter(shot[iSelPos1]);
			bool pos2IsOnDelimiter       = IsTrailingDelimiter(shot[iSelPos2 - 1]);
			bool areOnMatchingDelimiters = pos1IsOnDelimiter && pos2IsOnDelimiter && IsMatchingDelimiters(shot[iSelPos1], shot[iSelPos2 - 1]);
			int startAdjust = areOnMatchingDelimiters ? 1 : 0;
			int iStart1 = iSelPos1 - (canMovePos1 ? startAdjust : 0);
			int iStart2 = iSelPos2 + (canMovePos2 ? startAdjust : 0);
			for (int tstPos1 = iStart1; tstPos1 >= 0; --tstPos1)
			{
				char ch1 = shot[tstPos1];
				if (!IsLeadingDelimiter(ch1)) { continue; }
				for (int tstPos2 = iStart2; tstPos2 < selLen; ++tstPos2)
				{
					char ch2 = shot[tstPos2];
					if (!IsMatchingDelimiters(ch1, ch2)) { continue; }
					// For Select(); pos 1 is the position of the first char, but pos2 is the position _after_ the last char.
					if ((bothJustInsideDelimiter && tstPos1 != iStart1) ||
					    (pos1IsOnDelimiter && tstPos1 == iStart1))
					{
						++tstPos2; // include the delimiters
					}
					else
					{
						++tstPos1; // exclude the delimiters
					}
					textView.Selection.Select(new SnapshotSpan(new SnapshotPoint(shot, tstPos1), new SnapshotPoint(shot, tstPos2)), false);
					return VSConstants.S_OK;
				}
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
