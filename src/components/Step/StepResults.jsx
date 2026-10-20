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

const Field = ({ label, value }) => (
    <div className="doc-field">
        <span className="doc-field-label">{label}</span>
        <span className="doc-field-value">
            {value || '—'}
        </span>
    </div>
);

const formatDate = (iso) => {
    if (!iso) {
        return '—';
    }

    const [year, month, day] = iso.split('-');

    return `${day}.${month}.${year}`;
};

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
            return '—';
        }

        return value;
    };

    const materialName =
        formData.materialType === 'cardboard'
            ? 'Картон'
            : 'Бумага';

    const calcDate = new Date().toLocaleDateString('ru-RU');

    return (
        <div className="page active results-page">
            <div className="result-document">
                <header className="doc-head">
                    <div className="doc-brand">
                        <span className="doc-mark">КП</span>

                        <div className="doc-brand-text">
                            <strong>Калькулятор печати</strong>
                            <span>
                                Расчёт стоимости полиграфической продукции
                            </span>
                        </div>
                    </div>

                    <div className="doc-meta">
                        <div className="doc-meta-row">
                            <span>Документ</span>
                            <strong>Коммерческое предложение</strong>
                        </div>

                        <div className="doc-meta-row">
                            <span>Дата расчёта</span>
                            <strong>{calcDate}</strong>
                        </div>

                        <div className="doc-meta-row">
                            <span>Дата отгрузки</span>
                            <strong>
                                {formatDate(formData.shippingDate)}
                            </strong>
                        </div>

                        <div className="doc-status">
                            Расчёт выполнен
                        </div>
                    </div>
                </header>

                <div className="doc-rule" />

                <section className="doc-section">
                    <h3>Заказчик</h3>

                    <div className="doc-grid doc-grid--3">
                        <Field
                            label="Компания"
                            value={formData.companyName}
                        />

                        <Field
                            label="Адрес"
                            value={formData.companyAddress}
                        />

                        <Field
                            label="Контакты"
                            value={formData.companyContacts}
                        />
                    </div>
                </section>

                <section className="doc-section">
                    <h3>Изделие</h3>

                    <div className="doc-grid doc-grid--3">
                        <Field
                            label="Наименование"
                            value={formData.productName}
                        />

                        <Field
                            label="Тираж, шт"
                            value={formData.quantity}
                        />

                        <Field
                            label="Изделий на листе"
                            value={formData.perSheet}
                        />
                    </div>

                    {formData.notes && (
                        <p className="doc-note">
                            <span>Примечание:</span>
                            {formData.notes}
                        </p>
                    )}
                </section>

                <section className="doc-section">
                    <h3>Материал и формат</h3>

                    <div className="doc-grid doc-grid--3">
                        <Field
                            label="Материал"
                            value={materialName}
                        />

                        <Field
                            label="Цена материала"
                            value={`${formData.materialPrice || '—'} ${
                                formData.materialCurrency || '₽'
                            } / кг`}
                        />

                        <Field
                            label="Формат"
                            value={formData.formatSize}
                        />

                        <Field
                            label="Формат печати"
                            value={`${formData.printWidth || '—'} × ${
                                formData.printHeight || '—'
                            } мм`}
                        />

                        <Field
                            label="Закупочный формат"
                            value={`${formData.purchaseWidth || '—'} × ${
                                formData.purchaseHeight || '—'
                            } мм`}
                        />
                    </div>
                </section>

                <section className="doc-section">
                    <h3>Операции</h3>

                    <div className="doc-ops">
                        {operations.map(([label, field]) => (
                            <div
                                className="doc-op"
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
                </section>

                <section className="doc-section">
                    <h3>Расчётные показатели</h3>

                    <div className="doc-grid doc-grid--2">
                        <Field
                            label="Расчётный тираж"
                            value={results?.circulation}
                        />

                        <Field
                            label="Расход материала"
                            value={results?.sheetsKg}
                        />

                        <Field
                            label="Курс USD"
                            value={
                                results?.usdRate
                                    ? `${results.usdRate} ₽`
                                    : '—'
                            }
                        />

                        <Field
                            label="Курс EUR"
                            value={
                                results?.eurRate
                                    ? `${results.eurRate} ₽`
                                    : '—'
                            }
                        />
                    </div>
                </section>

                <section className="doc-totals">
                    <div className="doc-total-row">
                        <span>Итого</span>
                        <strong>{results?.total || '0.00 ₽'}</strong>
                    </div>

                    <div className="doc-total-row">
                        <span>НДС (20%)</span>
                        <strong>{results?.vat || '0.00 ₽'}</strong>
                    </div>

                    <div className="doc-total-row doc-total-row--final">
                        <span>Всего к оплате</span>
                        <strong>{results?.final || '0.00 ₽'}</strong>
                    </div>
                </section>

                <footer className="doc-foot">
                    <div className="doc-signs">
                        <div className="doc-sign">
                            <span>Исполнитель</span>
                            <em />
                        </div>

                        <div className="doc-sign">
                            <span>Заказчик</span>
                            <em />
                        </div>
                    </div>

                    <p className="doc-foot-note">
                        Расчёт выполнен автоматически · Калькулятор печати
                    </p>
                </footer>
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
                    🖨 Печать / Сохранить PDF
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
