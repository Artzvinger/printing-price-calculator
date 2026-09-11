import React from 'react';

const operations = [
    ['Подрезка', 'cuttingFormat'],
    ['Печать', 'printType'],
    ['Ламинация', 'lamination'],
    ['УФ-лак', 'uvVarnish'],
    ['Резка', 'cutting'],
    ['Тиснение 1', 'embossing1'],
    ['Тиснение 2', 'embossing2'],
    ['Вырубка', 'dieCutting'],
    ['Склейка', 'gluing'],
    ['Брошюровка', 'binding'],
];

const StepResults = ({
    formData,
    results,
    onBack,
    onClear,
    onNew,
}) => {
    const handlePrint = () => {
        window.print();
    };

    const formatOperation = (value) => {
        if (!value || value === 'нет') {
            return 'Нет';
        }

        return value;
    };

    const materialName =
        formData.materialType === 'cardboard'
            ? 'Картон'
            : 'Бумага';

    return (
        <div className="page active results-page">
            <div className="result-document">
                <div className="result-header">
                    <div>
                        <h2>Результат расчёта</h2>
                        <p className="result-subtitle">
                            Калькулятор печати
                        </p>
                    </div>

                    <div className="result-status">
                        Расчёт выполнен
                    </div>
                </div>

                <div className="result-block">
                    <h3>Информация о заказчике</h3>

                    <div className="result-grid">
                        <div className="result-field">
                            <span>Компания</span>
                            <strong>
                                {formData.companyName || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Адрес</span>
                            <strong>
                                {formData.companyAddress || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Контакты</span>
                            <strong>
                                {formData.companyContacts || '—'}
                            </strong>
                        </div>
                    </div>
                </div>

                <div className="result-block">
                    <h3>Информация об изделии</h3>

                    <div className="result-grid">
                        <div className="result-field">
                            <span>Наименование изделия</span>
                            <strong>
                                {formData.productName || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Количество изделий</span>
                            <strong>
                                {formData.quantity || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Изделий на листе</span>
                            <strong>
                                {formData.perSheet || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Дата отгрузки</span>
                            <strong>
                                {formData.shippingDate || '—'}
                            </strong>
                        </div>

                        {formData.notes && (
                            <div className="result-field result-field-wide">
                                <span>Примечания</span>
                                <strong>
                                    {formData.notes}
                                </strong>
                            </div>
                        )}
                    </div>
                </div>

                <div className="result-block">
                    <h3>Материал и формат</h3>

                    <div className="result-grid">
                        <div className="result-field">
                            <span>Материал</span>
                            <strong>{materialName}</strong>
                        </div>

                        <div className="result-field">
                            <span>Цена материала</span>
                            <strong>
                                {formData.materialPrice || '—'}{' '}
                                {formData.materialCurrency || '₽'} / кг
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Формат</span>
                            <strong>
                                {formData.formatSize || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Формат печати</span>
                            <strong>
                                {formData.printWidth || '—'} ×{' '}
                                {formData.printHeight || '—'} мм
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Закупочный формат</span>
                            <strong>
                                {formData.purchaseWidth || '—'} ×{' '}
                                {formData.purchaseHeight || '—'} мм
                            </strong>
                        </div>
                    </div>
                </div>

                <div className="result-block">
                    <h3>Операции</h3>

                    <div className="result-operations">
                        {operations.map(([label, field]) => (
                            <div
                                className="result-operation"
                                key={field}
                            >
                                <span>{label}</span>

                                <strong>
                                    {formatOperation(
                                        formData[field]
                                    )}
                                </strong>
                            </div>
                        ))}
                    </div>
                </div>

                <div className="result-block">
                    <h3>Расчётные показатели</h3>

                    <div className="result-grid">
                        <div className="result-field">
                            <span>Расчётный тираж</span>
                            <strong>
                                {results?.circulation || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Расход материала</span>
                            <strong>
                                {results?.sheetsKg || '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Курс USD</span>
                            <strong>
                                {results?.usdRate
                                    ? `${results.usdRate} ₽`
                                    : '—'}
                            </strong>
                        </div>

                        <div className="result-field">
                            <span>Курс EUR</span>
                            <strong>
                                {results?.eurRate
                                    ? `${results.eurRate} ₽`
                                    : '—'}
                            </strong>
                        </div>
                    </div>
                </div>

                <div className="result-price">
                    <div className="result-price-row">
                        <span>Итого:</span>
                        <strong>
                            {results?.total || '0.00 ₽'}
                        </strong>
                    </div>

                    <div className="result-price-row">
                        <span>НДС (20%):</span>
                        <strong>
                            {results?.vat || '0.00 ₽'}
                        </strong>
                    </div>

                    <div className="result-price-row result-price-final">
                        <span>Всего к оплате:</span>
                        <strong>
                            {results?.final || '0.00 ₽'}
                        </strong>
                    </div>
                </div>

                <div className="result-footer">
                    <span>
                        Расчёт выполнен автоматически
                    </span>

                    <span>
                        Калькулятор печати
                    </span>
                </div>
            </div>

            <div className="result-actions">
                <button
                    className="prev-btn"
                    onClick={onBack}
                    type="button"
                >
                    ← Назад
                </button>

                <button
                    className="print-btn"
                    onClick={handlePrint}
                    type="button"
                >
                    🖨 Печать / PDF
                </button>

                <button
                    id="clear-btn"
                    onClick={onClear}
                    type="button"
                >
                    Очистить данные
                </button>

                <button
                    id="new-calculation-btn"
                    onClick={onNew}
                    type="button"
                >
                    Новый расчёт
                </button>
            </div>
        </div>
    );
};

export default StepResults;