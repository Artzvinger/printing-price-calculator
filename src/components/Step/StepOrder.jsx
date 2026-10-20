import React from 'react';

const StepOrder = ({
                       formData,
                       updateFormData,
                       onCalculate,
                       onBack,
                       isLoading,
                       onRefreshRates,
                       isRatesLoading,
                       ratesStatus,
                   }) => {
    const handleChange = (e) => {
        const { id, value } = e.target;

        updateFormData(id, value);
    };

    return (
        <div className="page active">
            <div className="section">
                <h2>Информация о заказе</h2>

                <div className="order-info">
                    <input
                        type="text"
                        id="productName"
                        placeholder="Наименование изделия"
                        value={formData.productName || ''}
                        onChange={handleChange}
                    />

                    <input
                        type="number"
                        id="quantity"
                        placeholder="Количество изделий"
                        min="1"
                        value={formData.quantity || ''}
                        onChange={handleChange}
                    />

                    <input
                        type="number"
                        id="perSheet"
                        placeholder="Количество на листе"
                        step="0.1"
                        min="0.1"
                        value={formData.perSheet || ''}
                        onChange={handleChange}
                    />

                    <input
                        type="text"
                        id="notes"
                        placeholder="Примечания"
                        value={formData.notes || ''}
                        onChange={handleChange}
                    />
                </div>
            </div>

            <div className="section">
                <h2>Курсы валют (ЦБ РФ)</h2>

                <div className="rates-grid">
                    <div className="rate-field">
                        <label htmlFor="usdRate">USD:</label>

                        <div className="rate-input">
                            <input
                                type="number"
                                id="usdRate"
                                step="0.01"
                                min="0"
                                placeholder="Курс USD"
                                value={formData.usdRate || ''}
                                onChange={handleChange}
                            />

                            <span>₽</span>
                        </div>
                    </div>

                    <div className="rate-field">
                        <label htmlFor="eurRate">EUR:</label>

                        <div className="rate-input">
                            <input
                                type="number"
                                id="eurRate"
                                step="0.01"
                                min="0"
                                placeholder="Курс EUR"
                                value={formData.eurRate || ''}
                                onChange={handleChange}
                            />

                            <span>₽</span>
                        </div>
                    </div>

                    <button
                        type="button"
                        className="refresh-rates-btn"
                        onClick={onRefreshRates}
                        disabled={isRatesLoading}
                    >
                        {isRatesLoading
                            ? '⏳ Обновление...'
                            : '🔄 Обновить'}
                    </button>
                </div>

                {ratesStatus && (
                    <p className="rates-status">{ratesStatus}</p>
                )}
            </div>

            <div className="page-navigation">
                <button
                    className="prev-btn"
                    onClick={onBack}
                    type="button"
                    disabled={isLoading}
                >
                    ← Назад
                </button>

                <button
                    className="next-btn"
                    onClick={onCalculate}
                    type="button"
                    disabled={isLoading}
                >
                    {isLoading
                        ? '⏳ Расчёт...'
                        : 'Рассчитать →'}
                </button>
            </div>
        </div>
    );
};

export default StepOrder;