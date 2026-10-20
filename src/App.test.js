import { render, screen, fireEvent, act } from '@testing-library/react';
import App from './App';
import { fetchCurrencyRates } from './api/currencyApi';

jest.mock('./api/currencyApi');

describe('App', () => {
    beforeEach(() => {
        localStorage.clear();
        fetchCurrencyRates.mockResolvedValue({
            usdRate: 82.85,
            eurRate: 93.23,
        });
    });

    const renderApp = async () => {
        render(<App />);

        // дожидаемся асинхронной загрузки курсов валют
        await act(async () => {});
    };

    it('отображает заголовок и первый шаг (заказчик)', async () => {
        await renderApp();

        expect(
            screen.getByText('КАЛЬКУЛЯТОР ПЕЧАТИ')
        ).toBeInTheDocument();
        expect(
            screen.getByText('Информация о заказчике')
        ).toBeInTheDocument();
    });

    it('переходит на следующий шаг по кнопке «Далее»', async () => {
        await renderApp();

        fireEvent.click(screen.getByText('Далее →'));

        expect(
            screen.getByText('Выбор материала')
        ).toBeInTheDocument();
    });

    it('переключается по шагам через навигацию', async () => {
        await renderApp();

        fireEvent.click(screen.getByRole('button', { name: 'Материал' }));

        expect(
            screen.getByText('Выбор материала')
        ).toBeInTheDocument();
    });
});
