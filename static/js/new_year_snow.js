// static/js/new_year_snow.js
(function () {
    'use strict';

    // ═══════════════════════════════════════════════════════════
    //  НАЛАШТУВАННЯ
    // ═══════════════════════════════════════════════════════════
    const FLAKES_COUNT = 220;          // кількість сніжинок
    const MAX_SIZE = 5;                // максимальний радіус сніжинки (px)
    const MIN_SIZE = 2;                // мінімальний радіус
    const SPEED_Y_MIN = 2.7;           // вертикальна швидкість (px за кадр)
    const SPEED_Y_MAX = 3.2;
    const SPEED_X_AMPLITUDE = 0.5;     // амплітуда горизонтального погойдування
    const OPACITY_MIN = 0.35;
    const OPACITY_MAX = 0.85;

    // ═══════════════════════════════════════════════════════════
    //  ПЕРІОД ПОКАЗУ
    //  Сніг показується з 1 грудня по 14 січня включно.
    //  Змініть ці числа, якщо потрібен інший період.
    // ═══════════════════════════════════════════════════════════
    const START_MONTH = 12;  // грудень
    const START_DAY = 1;
    const END_MONTH = 1;     // січень
    const END_DAY = 14;

    function isNewYearPeriod() {
        const now = new Date();
        const month = now.getMonth() + 1; // 1–12
        const day = now.getDate();

        // Грудень (після START_DAY) АБО січень (до END_DAY)
        if (month === START_MONTH && day >= START_DAY) return true;
        if (month === END_MONTH && day <= END_DAY) return true;
        return false;
    }

    if (!isNewYearPeriod()) {
        return; // не святковий період — нічого не робимо
    }

    // ═══════════════════════════════════════════════════════════
    //  СТВОРЕННЯ CANVAS
    // ═══════════════════════════════════════════════════════════
    const canvas = document.createElement('canvas');
    canvas.id = 'new-year-snow-canvas';
    canvas.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        pointer-events: none;   /* кліки проходять крізь сніг */
        z-index: 0;             /* позаду контенту */
    `;
    document.body.insertBefore(canvas, document.body.firstChild);

    const ctx = canvas.getContext('2d');
    let width, height;
    let snowflakes = [];

    function resizeCanvas() {
        width = canvas.width = window.innerWidth;
        height = canvas.height = window.innerHeight;
    }

    // ═══════════════════════════════════════════════════════════
    //  СТВОРЕННЯ СНІЖИНОК
    // ═══════════════════════════════════════════════════════════
    function createSnowflake() {
        return {
            x: Math.random() * width,
            y: Math.random() * height - height, // починаємо вище екрана
            radius: MIN_SIZE + Math.random() * (MAX_SIZE - MIN_SIZE),
            speedY: SPEED_Y_MIN + Math.random() * (SPEED_Y_MAX - SPEED_Y_MIN),
            phase: Math.random() * Math.PI * 2,        // фаза погойдування
            phaseSpeed: 0.01 + Math.random() * 0.02,   // швидкість погойдування
            opacity: OPACITY_MIN + Math.random() * (OPACITY_MAX - OPACITY_MIN),
        };
    }

    function initSnowflakes() {
        snowflakes = [];
        for (let i = 0; i < FLAKES_COUNT; i++) {
            snowflakes.push(createSnowflake());
        }
    }

    // ═══════════════════════════════════════════════════════════
    //  АНІМАЦІЯ
    // ═══════════════════════════════════════════════════════════
    function animate() {
        ctx.clearRect(0, 0, width, height);

        for (const flake of snowflakes) {
            // Рух вниз
            flake.y += flake.speedY;

            // Горизонтальне погойдування (синусоїда)
            flake.phase += flake.phaseSpeed;
            const offsetX = Math.sin(flake.phase) * SPEED_X_AMPLITUDE;

            // Якщо сніжинка вийшла за нижню межу — повертаємо нагору
            if (flake.y > height + 10) {
                flake.y = -10;
                flake.x = Math.random() * width;
            }

            // Малювання
            ctx.beginPath();
            ctx.arc(flake.x + offsetX, flake.y, flake.radius, 0, Math.PI * 2);
            ctx.fillStyle = `rgba(255, 255, 255, ${flake.opacity})`;
            ctx.fill();
        }

        requestAnimationFrame(animate);
    }

    // ═══════════════════════════════════════════════════════════
    //  ЗАПУСК
    // ═══════════════════════════════════════════════════════════
    window.addEventListener('resize', () => {
        resizeCanvas();
        // Після зміни розміру вікна сніжинки можуть опинитися за межами —
        // переініціалізуємо їх, щоб усе виглядало коректно.
        initSnowflakes();
    });

    resizeCanvas();
    initSnowflakes();
    animate();
})();