const { createProxyMiddleware } = require('http-proxy-middleware');

// URL развёрнутого веб-приложения Google Apps Script.
// Переопределяется переменной окружения GAS_WEB_APP_URL (см. .env.example).
const GAS_WEB_APP_URL =
    process.env.GAS_WEB_APP_URL ||
    'https://script.google.com/macros/s/AKfycbzYAbY4S0quoK-lLgGiIjlfmzDwsnTVjuK8_1qlkycGvzk5BdvmELmFPS7QDWheWTSo/exec';

const target = new URL(GAS_WEB_APP_URL);

module.exports = function (app) {
    app.use(
        '/api/proxy',
        createProxyMiddleware({
            target: target.origin,
            changeOrigin: true,
            pathRewrite: () => target.pathname,
        })
    );
};
