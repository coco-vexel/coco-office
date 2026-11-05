/** Framework-free counterpart of the hosts' anchored destructive-action popover. */
export class ActionConfirmPopover {
  private readonly element = document.createElement('div');
  private anchor: HTMLButtonElement | null = null;

  constructor(private readonly host: HTMLElement, root: HTMLElement | ShadowRoot) {
    this.element.setAttribute('role', 'dialog');
    Object.assign(this.element.style, {
      position: 'absolute', zIndex: '60', display: 'none', width: '256px',
      maxWidth: 'calc(100% - 8px)', boxSizing: 'border-box', padding: '12px',
      border: '1px solid hsl(var(--border, 240 5.9% 90%))', borderRadius: '8px',
      background: 'hsl(var(--popover, 0 0% 100%))', color: 'hsl(var(--popover-foreground, 225 13% 10%))',
      boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1)',
      font: '13px Arial, sans-serif',
    });
    for (const name of ['pointerdown', 'click', 'dblclick', 'contextmenu']) {
      this.element.addEventListener(name, (event) => event.stopPropagation());
    }
    this.element.addEventListener('keydown', (event) => {
      event.stopPropagation();
      if (event.key === 'Escape') {
        event.preventDefault();
        this.hide();
      }
    });
    root.appendChild(this.element);
  }

  show(anchor: HTMLButtonElement, options: { title: string; description: string; onConfirm: () => void }): void {
    if (this.anchor === anchor) {
      this.hide();
      return;
    }
    this.hide(false);
    this.anchor = anchor;
    anchor.setAttribute('aria-expanded', 'true');
    anchor.setAttribute('aria-haspopup', 'dialog');
    this.element.setAttribute('aria-label', options.title);
    const title = document.createElement('div');
    title.textContent = options.title;
    Object.assign(title.style, { fontSize: '14px', fontWeight: '600', lineHeight: '20px' });
    const description = document.createElement('div');
    description.textContent = options.description;
    description.title = options.description;
    Object.assign(description.style, {
      marginTop: '4px', fontSize: '12px', lineHeight: '18px',
      color: 'hsl(var(--muted-foreground, 240 3.8% 46.1%))', overflowWrap: 'anywhere',
      display: '-webkit-box', WebkitBoxOrient: 'vertical', WebkitLineClamp: '2', overflow: 'hidden',
    });
    const warning = document.createElement('div');
    warning.textContent = '此操作无法撤销。';
    Object.assign(warning.style, {
      marginTop: '4px', fontSize: '12px', lineHeight: '18px',
      color: 'hsl(var(--muted-foreground, 240 3.8% 46.1%))',
    });
    const actions = document.createElement('div');
    Object.assign(actions.style, { display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '12px' });
    const cancel = this.button('取消', false);
    cancel.addEventListener('click', () => this.hide());
    const confirm = this.button('删除', true);
    confirm.addEventListener('click', () => {
      if (!this.anchor) return;
      this.hide(false);
      options.onConfirm();
    });
    actions.append(cancel, confirm);
    this.element.replaceChildren(title, description, warning, actions);
    this.element.style.display = 'block';
    this.position();
    document.addEventListener('pointerdown', this.handleOutsideInteraction, true);
    document.addEventListener('focusin', this.handleOutsideInteraction, true);
    window.addEventListener('resize', this.handleResize);
    cancel.focus({ preventScroll: true });
  }

  private position(): void {
    if (!this.anchor) return;
    const hostRect = this.host.getBoundingClientRect();
    const anchorRect = this.anchor.getBoundingClientRect();
    const bounds = this.element.getBoundingClientRect();
    this.element.style.left = `${Math.max(4, Math.min(anchorRect.left - hostRect.left, hostRect.width - bounds.width - 4))}px`;
    this.element.style.top = `${Math.max(4, anchorRect.top - hostRect.top - bounds.height - 6)}px`;
  }

  hide(restoreFocus = true): void {
    const anchor = this.anchor;
    this.anchor = null;
    this.element.style.display = 'none';
    anchor?.removeAttribute('aria-expanded');
    document.removeEventListener('pointerdown', this.handleOutsideInteraction, true);
    document.removeEventListener('focusin', this.handleOutsideInteraction, true);
    window.removeEventListener('resize', this.handleResize);
    if (restoreFocus && anchor?.isConnected) anchor.focus({ preventScroll: true });
  }

  destroy(): void {
    this.hide(false);
    this.element.remove();
  }

  private readonly handleOutsideInteraction = (event: Event): void => {
    const path = event.composedPath();
    if (!path.includes(this.element) && (!this.anchor || !path.includes(this.anchor))) this.hide(false);
  };

  private readonly handleResize = (): void => { this.position(); };

  private button(label: string, destructive: boolean): HTMLButtonElement {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = label;
    Object.assign(button.style, {
      height: '32px', padding: '0 12px', borderRadius: '8px', cursor: 'pointer',
      font: '500 12px Arial, sans-serif', boxSizing: 'border-box',
      border: destructive ? '1px solid transparent' : '1px solid hsl(var(--input, 240 5.9% 90%))',
      background: destructive ? 'hsl(var(--destructive, 0 84.2% 60.2%))' : 'hsl(var(--background, 0 0% 100%))',
      color: destructive ? 'hsl(var(--destructive-foreground, 0 0% 98%))' : 'inherit',
      boxShadow: '0 1px 2px 0 rgb(0 0 0 / 0.05)',
    });
    button.addEventListener('mouseenter', () => { button.style.filter = 'brightness(0.96)'; });
    button.addEventListener('mouseleave', () => { button.style.filter = ''; });
    return button;
  }
}
