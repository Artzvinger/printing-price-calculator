import React from 'react';

const StepOrder = ({
                       formData,
                       updateFormData,
                       onCalculate,
                       onBack,
                       isLoading,
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