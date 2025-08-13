const DEFAULT_DURATION = 8000;
let toastCounter = 0;

function ensureContainer() {
    let container = document.getElementById('toast-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        container.setAttribute('role', 'region');
        container.setAttribute('aria-label', 'Notifications');
        document.body.appendChild(container);
    }
    return container;
}

function typeMeta(type) {
    switch (type) {
        case 'info': return { icon: 'ℹ️', ariaRole: 'status' };
        case 'warning': return { icon: '⚠️', ariaRole: 'alert' };
        case 'error': return { icon: '❌', ariaRole: 'alert' };
        case 'success':
        default: return { icon: '✅', ariaRole: 'status' };
    }
}

export function showToast(message, typeOrOptions, extraOptions) {
    // Argument normalization
    let options = {};
    if (typeof typeOrOptions === 'string') {
        options.type = typeOrOptions;
        if (extraOptions && typeof extraOptions === 'object') {
            options = { ...extraOptions, type: typeOrOptions };
        }
    } else if (typeof typeOrOptions === 'object' && typeOrOptions !== null) {
        options = { ...typeOrOptions };
    }

    const {
        type = 'success',
        duration = DEFAULT_DURATION,
        dismissible = true,
        allowHTML = false,
        icon, // optional override
    } = options;

    const { icon: defaultIcon, ariaRole } = typeMeta(type);
    const finalIcon = icon === false ? '' : (icon || defaultIcon);

    const container = ensureContainer();

    // Legacy single #toast support (fallback if container CSS not loaded yet)
    const legacy = document.getElementById('toast');
    if (!container && legacy) {
        legacy.textContent = message;
        legacy.style.display = 'block';
        setTimeout(() => { legacy.style.display = 'none'; }, duration);
        return;
    }

    const id = `toast-${Date.now()}-${++toastCounter}`;
    const el = document.createElement('div');
    el.className = `toast toast--${type}`;
    el.id = id;
    el.setAttribute('role', ariaRole);
    el.setAttribute('aria-live', ariaRole === 'alert' ? 'assertive' : 'polite');

    const content = document.createElement('div');
    content.className = 'toast-content';
    if (finalIcon) {
        const iconSpan = document.createElement('span');
        iconSpan.className = 'toast-icon';
        iconSpan.textContent = finalIcon;
        content.appendChild(iconSpan);
    }
    const msgSpan = document.createElement('span');
    msgSpan.className = 'toast-message';
    if (allowHTML) {
        msgSpan.innerHTML = message;
    } else {
        msgSpan.textContent = message;
    }
    content.appendChild(msgSpan);
    el.appendChild(content);

    if (dismissible) {
        const closeBtn = document.createElement('button');
        closeBtn.className = 'toast-close';
        closeBtn.setAttribute('aria-label', 'Dismiss notification');
        closeBtn.innerHTML = '&times;';
        closeBtn.addEventListener('click', () => dismissToast(el));
        el.appendChild(closeBtn);
    }

    // Progress bar
    if (duration > 0) {
        const progressWrap = document.createElement('div');
        progressWrap.className = 'toast-progress';
        const bar = document.createElement('div');
        bar.className = 'toast-progress-bar';
        progressWrap.appendChild(bar);
        el.appendChild(progressWrap);
    }

    container.appendChild(el);

    // Allow animation frame for CSS enter animation.
    requestAnimationFrame(() => {
        el.classList.add('toast-enter');
    });

    let remaining = duration;
    let startTime = performance.now();
    let timerId;
    let rafId;
    const progressBar = el.querySelector('.toast-progress-bar');

    function tick(now) {
        if (!progressBar) return;
        const elapsed = now - startTime;
        const pct = Math.max(0, 1 - (elapsed / duration));
        progressBar.style.transform = `scaleX(${pct})`;
        if (elapsed >= duration) return;
        rafId = requestAnimationFrame(tick);
    }

    function startTimers() {
        if (duration > 0) {
            timerId = setTimeout(() => dismissToast(el), remaining);
            if (progressBar) {
                startTime = performance.now();
                rafId = requestAnimationFrame(tick);
            }
        }
    }

    function pauseTimers() {
        if (timerId) {
            clearTimeout(timerId);
            timerId = null;
        }
        if (progressBar && rafId) {
            cancelAnimationFrame(rafId);
            rafId = null;
            const computed = progressBar.style.transform;
            // Extract scaleX value
            const match = /scaleX\(([^)]+)\)/.exec(computed);
            if (match) {
                const pct = parseFloat(match[1]);
                remaining = pct * duration;
            }
        }
    }

    el.addEventListener('mouseenter', pauseTimers);
    el.addEventListener('mouseleave', startTimers);

    startTimers();
    return id;
}

export function dismissToast(elOrId) {
    const el = typeof elOrId === 'string' ? document.getElementById(elOrId) : elOrId;
    if (!el) return;
    el.classList.add('toast-leave');
    el.addEventListener('animationend', () => {
        if (el.parentElement) el.parentElement.removeChild(el);
    });
}

export function clearToasts() {
    const container = document.getElementById('toast-container');
    if (!container) return;
    Array.from(container.children).forEach(ch => dismissToast(ch));
}
