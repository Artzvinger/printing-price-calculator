const API_URL = '/api/proxy';

export const calculateOrder = async (formData) => {
    const response = await fetch(API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData),
    });

    const text = await response.text();
    try {
        return JSON.parse(text);
    } catch {
        throw new Error('Некорректный ответ от сервера');
    }
};