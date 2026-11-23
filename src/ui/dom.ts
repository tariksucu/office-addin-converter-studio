/**
 * A few DOM helpers. Text always goes in as text (never as HTML): file names and the API's
 * messages are shown as they are, and nothing from outside can become markup.
 */

type Child = Node | string | null | undefined | false;
type Props = Record<string, string | number | boolean | null | undefined | EventListener>;

export function h<K extends keyof HTMLElementTagNameMap>(
  doc: Document,
  tag: K,
  props: Props = {},
  ...children: Child[]
): HTMLElementTagNameMap[K] {
  const element = doc.createElement(tag);
  for (const [name, value] of Object.entries(props)) {
    if (value === null || value === undefined || value === false) continue;
    if (name.startsWith('on') && typeof value === 'function')
      element.addEventListener(name.slice(2).toLowerCase(), value);
    else if (name === 'className') element.className = String(value);
    else if (value === true) element.setAttribute(name, '');
    else element.setAttribute(name, String(value));
  }
  append(element, ...children);
  return element;
}

export function append(parent: Node, ...children: Child[]): void {
  for (const child of children) {
    if (child === null || child === undefined || child === false) continue;
    parent.appendChild(
      typeof child === 'string' ? parent.ownerDocument!.createTextNode(child) : child,
    );
  }
}

export function clear(element: Element): void {
  while (element.firstChild) element.removeChild(element.firstChild);
}

let counter = 0;
/** A unique id for a label's `for`. */
export function uid(prefix: string): string {
  counter += 1;
  return `${prefix}-${counter}`;
}
