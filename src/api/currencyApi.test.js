import { fetchCurrencyRates } from './currencyApi';

describe('fetchCurrencyRates', () => {
    beforeEach(() => {
        global.fetch = jest.fn();
    });

    afterEach(() => {
        jest.restoreAllMocks();
    });

    it('возвращает округлённые курсы USD и EUR', async () => {
        global.fetch.mockResolvedValue({
            ok: true,
            json: async () => ({
                Valute: {
                    USD: { Value: 82.8504 },
                    EUR: { Value: 93.2316 },
                },
            }),
        });

        const rates = await fetchCurrencyRates();

        expect(rates).toEqual({
            usdRate: 82.85,
            eurRate: 93.23,
        });
        expect(global.fetch).toHaveBeenCalledTimes(1);
    });

    it('бросает ошибку при HTTP-ошибке ЦБ РФ', async () => {
        global.fetch.mockResolvedValue({
            ok: false,
            status: 500,
        });

        await expect(fetchCurrencyRates()).rejects.toThrow(
            'ЦБ РФ вернул ошибку: 500'
        );
    });

    it('бросает ошибку при некорректном ответе', async () => {
        global.fetch.mockResolvedValue({
            ok: true,
            json: async () => ({ Valute: {} }),
        });

        await expect(fetchCurrencyRates()).rejects.toThrow(
            'ЦБ РФ вернул некорректные данные'
        );
    });
});
