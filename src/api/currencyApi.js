const CBR_API_URL = 'https://www.cbr-xml-daily.ru/daily_json.js';

const round = (value, digits = 2) => {
    const factor = 10 ** digits;
    return Math.round(value * factor) / factor;
};

export const fetchCurrencyRates = async () => {
    const response = await fetch(CBR_API_URL);

    if (!response.ok) {
        throw new Error(`ЦБ РФ вернул ошибку: ${response.status}`);
    }

    const data = await response.json();

    const usd = data?.Valute?.USD?.Value;
    const eur = data?.Valute?.EUR?.Value;

    if (!usd || !eur) {
        throw new Error('ЦБ РФ вернул некорректные данные');
    }

    return {
        usdRate: round(usd),
        eurRate: round(eur),
    };
};
