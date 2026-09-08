import React from 'react';

const StepResults = ({
                         formData,
                         results,
                         onBack,
                         onClear,
                         onNew,
                     }) => {
    return (
        <div className="page active">
            <div className="section results">
                <h2>Результаты расчета</h2>

                <div className="result-item">
                    <span>Итого:</span>
                    <span id="total-result">
                        {results?.total || '0.00 ₽'}
                    </span>
                </div>

                <div className="result-item">
                    <span>НДС (20%):</span>
                    <span id="vat-result">
                        {results?.vat || '0.00 ₽'}
                    </span>
                </div>

                <div className="result-item total">
                    <span>Всего к оплате:</span>
                    <span id="final-result">
                        {results?.final || '0.00 ₽'}
                    </span>
                </div>
            </div>

            <div className="page-navigation">
                <button
                    className="prev-btn"
                    onClick={onBack}
                    type="button"
                >
                    ← Назад
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
                    Новый расчет
                </button>
            </div>
        </div>
    );
};

export default StepResults;