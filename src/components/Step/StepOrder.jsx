import React from 'react';

const StepOrder = ({ data, onChange, onNext, onPrev }) => {
    const handleChange = (e) => {
        onChange({ ...data, [e.target.id]: e.target.value });
    };

    return (
        <div className="page active">
            <div className="section">
                <h2>Информация о заказе</h2>
                <div className="order-info">
                    <input
                        type="text"
                        id="product-name"
                        placeholder="Наименование изделия"
                        value={data.productName || ''}
                        onChange={handleChange}
                    />
                    <input
                        type="number"
                        id="quantity"
                        placeholder="Количество изделий"
                        min="1"
                        value={data.quantity || ''}
                        onChange={handleChange}
                    />
                    <input
                        type="number"
                        id="per-sheet"
                        placeholder="Количество на листе"
                        step="0.1"
                        min="0.1"
                        value={data.perSheet || ''}
                        onChange={handleChange}
                    />
                    <input
                        type="text"
                        id="notes"
                        placeholder="Примечания"
                        value={data.notes || ''}
                        onChange={handleChange}
                    />
                </div>
            </div>
            <div className="page-navigation">
                <button className="prev-btn" onClick={onPrev}>← Назад</button>
                <button className="next-btn" onClick={onNext}>Далее →</button>
            </div>
        </div>
    );
};

export default StepOrder;