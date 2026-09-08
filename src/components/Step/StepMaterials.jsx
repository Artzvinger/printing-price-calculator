import React from 'react';

const StepMaterials = ({
                           formData,
                           updateFormData,
                           onNext,
                           onBack,
                       }) => {
    const handleChange = (e) => {
        const { id, value } = e.target;

        updateFormData(id, value);
    };

    return (
        <div className="page active">
            <div className="section">
                <h2>Выбор материала</h2>

                <div className="material-cost">
                    <select
                        id="materialType"
                        value={formData.materialType || 'paper'}
                        onChange={handleChange}
                    >
                        <option value="paper">Бумага</option>
                        <option value="cardboard">Картон</option>
                    </select>

                    <input
                        type="number"
                        id="materialPrice"
                        placeholder="Цена за кг"
                        value={formData.materialPrice || ''}
                        onChange={handleChange}
                        step="0.01"
                    />

                    <select
                        id="materialCurrency"
                        value={formData.materialCurrency || '₽'}
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
                    <label htmlFor="formatSize">
                        Формат бумаги:
                    </label>

                    <select
                        id="formatSize"
                        value={formData.formatSize || 'А3'}
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
                                id="printWidth"
                                value={formData.printWidth || ''}
                                onChange={handleChange}
                            />

                            ×

                            <input
                                type="number"
                                id="printHeight"
                                value={formData.printHeight || ''}
                                onChange={handleChange}
                            />

                            мм
                        </div>

                        <div className="format-compact">
                            <span>Закупочный формат:</span>

                            <input
                                type="number"
                                id="purchaseWidth"
                                value={formData.purchaseWidth || ''}
                                onChange={handleChange}
                            />

                            ×

                            <input
                                type="number"
                                id="purchaseHeight"
                                value={formData.purchaseHeight || ''}
                                onChange={handleChange}
                            />

                            мм
                        </div>
                    </div>
                </div>
            </div>

            <div className="page-navigation">
                <button
                    type="button"
                    className="prev-btn"
                    onClick={onBack}
                >
                    ← Назад
                </button>

                <button
                    type="button"
                    className="next-btn"
                    onClick={onNext}
                >
                    Далее →
                </button>
            </div>
        </div>
    );
};

export default StepMaterials;