import { render, screen, fireEvent } from '@testing-library/react';
import StepOrder from './StepOrder';

const buildProps = (overrides = {}) => ({
    formData: {},
    updateFormData: jest.fn(),
    onCalculate: jest.fn(),
    onBack: jest.fn(),
    isLoading: false,
    onRefreshRates: jest.fn(),
    isRatesLoading: false,
    ratesStatus: '',
    ...overrides,
});

describe('StepOrder', () => {
    it('отображает поля заказа и блок курсов валют', () => {
        render(<StepOrder {...buildProps()} />);

        expect(
            screen.getByPlaceholderText('Наименование изделия')
        ).toBeInTheDocument();
        expect(
            screen.getByPlaceholderText('Количество изделий')
        ).toBeInTheDocument();
        expect(
            screen.getByText('Курсы валют (ЦБ РФ)')
        ).toBeInTheDocument();
    });

    it('вызывает onCalculate при нажатии «Рассчитать»', () => {
        const props = buildProps();

        render(<StepOrder {...props} />);

        fireEvent.click(screen.getByText(/Рассчитать/));

        expect(props.onCalculate).toHaveBeenCalledTimes(1);
    });

    it('вызывает onRefreshRates при нажатии «Обновить»', () => {
        const props = buildProps();

        render(<StepOrder {...props} />);

        fireEvent.click(screen.getByText(/Обновить/));

        expect(props.onRefreshRates).toHaveBeenCalledTimes(1);
    });

    it('обновляет данные через updateFormData при вводе', () => {
        const props = buildProps();

        render(<StepOrder {...props} />);

        fireEvent.change(
            screen.getByPlaceholderText('Наименование изделия'),
            { target: { value: 'Визитки' } }
        );

        expect(props.updateFormData).toHaveBeenCalledWith(
            'productName',
            'Визитки'
        );
    });
});
