import { attr, childNodesOf, hasClass, isElement, tagNameOf } from '../util/dom.js';
import { parseInlineStyle } from './WordCssTokenizer.js';
import { DiagnosticCode } from '../diagnostics/UnsupportedFeature.js';
import type { WordParseContext } from './WordParseContext.js';

/**
 * Word review comments (annotations).
 *
 * When a document with comments is copied, Word appends the comment text to
 * the end of the payload in a container it marks with
 * `mso-element:comment-list`, and puts anchors in the body:
 *
 *     <a class=msocomanchor id="_anchor_1" href="#_msocom_1" ...>[JD1]</a>
 *     <div style='mso-element:comment-list'>
 *       <div style='mso-element:comment'>… the comment text …</div>
 *     </div>
 *
 * Rendering that container inline would append every comment to the bottom of
 * the pasted content as body text, which is wrong and surprising. Dropping it
 * without a word would lose it entirely. So the comment list is lifted out and
 * reported, and the anchors are removed from the flow.
 */

export interface WordComment {
  id: string;
  /** Reviewer initials Word put in the anchor, e.g. `JD1`. */
  reference?: string;
  /** Author, when Word wrote it as a `title`/`mso-comment-*` attribute. */
  author?: string;
  text: string;
}

export interface CommentExtraction {
  comments: WordComment[];
  /** Number of anchors removed from the content flow. */
  anchorsRemoved: number;
}

/**
 * Remove comment containers and anchors from the working tree, returning the
 * comment text so a consumer can present it separately.
 */
export function extractComments(root: Element, ctx: WordParseContext): CommentExtraction {
  const result: CommentExtraction = { comments: [], anchorsRemoved: 0 };
  const containers: Element[] = [];
  const anchors: Element[] = [];

  const walk = (node: Node): void => {
    for (const child of childNodesOf(node)) {
      if (!isElement(child)) continue;
      const style = parseInlineStyle(attr(child, 'style'));
      const element = (style['mso-element'] ?? '').toLowerCase();

      if (element === 'comment-list' || hasClass(child, 'msocomtxt')) {
        containers.push(child);
        continue;
      }
      if (
        hasClass(child, 'msocomanchor') ||
        hasClass(child, 'MsoCommentReference') ||
        element === 'comment' ||
        (tagNameOf(child) === 'a' && /^#_msocom_/i.test(attr(child, 'href') ?? ''))
      ) {
        anchors.push(child);
        continue;
      }
      walk(child);
    }
  };
  walk(root);

  for (const container of containers) {
    for (const comment of readComments(container)) result.comments.push(comment);
    container.parentNode?.removeChild(container);
  }
  for (const anchor of anchors) {
    anchor.parentNode?.removeChild(anchor);
    result.anchorsRemoved++;
  }

  if (result.comments.length > 0 || result.anchorsRemoved > 0) {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_UNSUPPORTED_FIELD,
      `${result.comments.length} Word review comment(s) were found. HTML has no comment primitive, so the anchors were removed from the content and the comment text is reported separately rather than appended to the document body.`,
      { details: { comments: result.comments.length, anchors: result.anchorsRemoved } },
    );
  }
  return result;
}

function readComments(container: Element): WordComment[] {
  const comments: WordComment[] = [];
  let index = 0;
  const walk = (node: Node): void => {
    for (const child of childNodesOf(node)) {
      if (!isElement(child)) continue;
      const style = parseInlineStyle(attr(child, 'style'));
      if ((style['mso-element'] ?? '').toLowerCase() === 'comment') {
        const text = (child.textContent ?? '').replace(/\s+/g, ' ').trim();
        if (text) {
          const comment: WordComment = { id: attr(child, 'id') ?? `comment-${++index}`, text };
          const reference = /^\[([^\]]+)\]/.exec(text)?.[1];
          if (reference) comment.reference = reference;
          comments.push(comment);
        }
        continue;
      }
      walk(child);
    }
  };
  walk(container);

  if (comments.length === 0) {
    const text = (container.textContent ?? '').replace(/\s+/g, ' ').trim();
    if (text) comments.push({ id: 'comment-1', text });
  }
  return comments;
}

/**
 * Word footnote and endnote containers use the same trick
 * (`mso-element:footnote-list`). They are lifted out for the same reason.
 */
export function extractNotes(root: Element, ctx: WordParseContext): string[] {
  const notes: string[] = [];
  const containers: Element[] = [];

  const walk = (node: Node): void => {
    for (const child of childNodesOf(node)) {
      if (!isElement(child)) continue;
      const style = parseInlineStyle(attr(child, 'style'));
      const element = (style['mso-element'] ?? '').toLowerCase();
      if (element === 'footnote-list' || element === 'endnote-list') {
        containers.push(child);
        continue;
      }
      walk(child);
    }
  };
  walk(root);

  for (const container of containers) {
    const text = (container.textContent ?? '').replace(/\s+/g, ' ').trim();
    if (text) notes.push(text);
    container.parentNode?.removeChild(container);
  }

  if (notes.length > 0) {
    ctx.diagnostics.warn(
      DiagnosticCode.WORD_FOOTNOTE_APPROXIMATED,
      `${notes.length} footnote/endnote container(s) were lifted out of the content flow and reported separately; HTML has no footnote primitive.`,
      { details: { notes: notes.length } },
    );
  }
  return notes;
}
