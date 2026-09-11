const API_URL = '/api/proxy';

export const calculateOrder = async (formData) => {
    try {
        const response = await fetch(API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(formData),
        });

        const text = await response.text();

        let data;

        try {
            data = text ? JSON.parse(text) : null;
        } catch {
            throw new Error(
                `Сервер вернул некорректный ответ: ${
    text || 'пустой ответ'
}`
            );
        }

        if (!response.ok) {
            throw new Error(
                data?.error ||
                data?.message ||
                `Ошибка сервера: ${response.status}`
            );
        }

        if (!data) {
            throw new Error('Сервер вернул пустой ответ');
        }

        if (!data.success) {
            throw new Error(
                data.error ||
                data.message ||
                'Не удалось выполнить расчёт'
            );
        }

        return data;
    } catch (error) {
        console.error('Ошибка расчёта:', error);

        if (error instanceof TypeError) {
            throw new Error(
                'Не удалось подключиться к серверу'
            );
        }

        throw error;
    }
};