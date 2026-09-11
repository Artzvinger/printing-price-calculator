export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    if (req.method !== 'POST') {
        return res.status(405).json({
            success: false,
            error: 'Метод не поддерживается',
        });
    }

    const TARGET_URL =
        'https://script.google.com/macros/s/AKfycbzYAbY4S0quoK-lLgGiIjlfmzDwsnTVjuK8_1qlkycGvzk5BdvmELmFPS7QDWheWTSo/exec';

    try {
        const response = await fetch(TARGET_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(req.body),
        });

        const text = await response.text();

        let data;

        try {
            data = text ? JSON.parse(text) : null;
        } catch {
            return res.status(502).json({
                success: false,
                error: 'Google Apps Script вернул некорректный JSON',
                response: text,
            });
        }

        if (!data) {
            return res.status(502).json({
                success: false,
                error: 'Google Apps Script вернул пустой ответ',
            });
        }

        return res
            .status(response.ok ? 200 : response.status)
            .json(data);
    } catch (error) {
        console.error('Proxy error:', error);

        return res.status(500).json({
            success: false,
            error: 'Ошибка соединения с Google Apps Script',
            details:
                error instanceof Error
                    ? error.message
                    : 'Неизвестная ошибка',
        });
    }
}