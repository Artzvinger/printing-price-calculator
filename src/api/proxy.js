export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    const TARGET_URL = 'https://script.google.com/macros/s/AKfycbzYAbY4S0quoK-lLgGiIjlfmzDwsnTVjuK8_1qlkycGvzk5BdvmELmFPS7QDWheWTSo/exec';

    try {
        const response = await fetch(TARGET_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(req.body),
        });

        const data = await response.json();
        res.status(200).json(data);
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
}