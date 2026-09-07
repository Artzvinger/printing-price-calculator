import React from 'react';

const StepMaterials = ({ data, onChange, onNext, onPrev }) => {
    const handleChange = (e) => {
        onChange({ ...data, [e.target.id]: e.target.value });
    };

    return (
        <div className="page active">
            <div className="section">
                <h2>Выбор материала</h2>
                <div className="material-cost">
                    <select
                        id="material-type-select"
                        value={data.materialType || 'paper'}
                        onChange={handleChange}
                    >
                        <option value="paper">Бумага</option>
                        <option value="cardboard">Картон</option>
                    </select>
                    <input
                        type="number"
                        id="material-price"
                        placeholder="Цена за кг"
                        value={data.materialPrice || ''}
                        onChange={handleChange}
                        step="0.01"
                    />
                    <select
                        id="material-currency"
                        value={data.materialCurrency || '₽'}
                        onChange={handleChange}
                    >
                        <option value="₽">₽</option>
                        <option value="$">$</option>
                        <option value="€">€</option>
                    </select>
                </div>
            </div>

            <div className="section">
                <h2>Информация о материале</h2>
                <div className="material-info">
                    <label htmlFor="format-size">Формат бумаги:</label>
                    <select
                        id="format-size"
                        value={data.formatSize || 'А3'}
                        onChange={handleChange}
                    >
                        <option value="А2">А2</option>
                        <option value="А3">А3</option>
                        <option value="А4">А4</option>
                    </select>

                    <div className="format-group">
                        <h3>Форматы материала</h3>
                        <div className="format-compact">
                            <span>Формат печати:</span>
                            <input
                                type="number"
                                id="print-width"
                                value={data.printWidth || ''}
                                onChange={handleChange}
                            />
                            ×
                            <input
                                type="number"
                                id="print-height"
                                value={data.printHeight || ''}
                                onChange={handleChange}
                            /> мм
                        </div>
                        <div className="format-compact">
                            <span>Закупочный формат:</span>
                            <input
                                type="number"
                                id="purchase-width"
                                value={data.purchaseWidth || ''}
                                onChange={handleChange}
                            />
                            ×
                            <input
                                type="number"
                                id="purchase-height"
                                value={data.purchaseHeight || ''}
                                onChange={handleChange}
                            /> мм
                        </div>
                    </div>
                </div>
            </div>

            <div className="page-navigation">
                <button className="prev-btn" onClick={onPrev}>← Назад</button>
                <button className="next-btn" onClick={onNext}>Далее →</button>
            </div>
        </div>
    );
};

export default StepMaterials;