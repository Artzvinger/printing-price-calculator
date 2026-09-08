import React from 'react';

const operations = [
    { id: 'cuttingFormat', label: 'Подрезка:' },
    { id: 'printType', label: 'Печать:' },
    { id: 'lamination', label: 'Ламинация:' },
    { id: 'uvVarnish', label: 'УФ-лак:' },
    { id: 'cutting', label: 'Резка:' },
    { id: 'embossing1', label: 'Тиснение (1):' },
    { id: 'embossing2', label: 'Тиснение (2):' },
    { id: 'dieCutting', label: 'Вырубка:' },
    { id: 'gluing', label: 'Склейка:' },
    { id: 'binding', label: 'Брошюровка:' },
];

const StepOperations = ({
                            formData,
                            updateFormData,
                            onNext,
                            onBack,
                        }) => {
    const handleChange = (e) => {
        const { id, value } = e.target;

        updateFormData(id, value);
    };

    const getOptions = (id) => {
        if (id === 'cuttingFormat') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="на формат А2">На формат А2</option>
                    <option value="на формат А3">На формат А3</option>
                    <option value="на формат А4">На формат А4</option>
                </>
            );
        }

        if (id === 'printType') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="1+0">1+0</option>
                    <option value="2+0">2+0</option>
                    <option value="3+0">3+0</option>
                    <option value="4+0">4+0</option>
                    <option value="5+0">5+0</option>
                </>
            );
        }

        if (id === 'lamination' || id === 'uvVarnish') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="Глянцевая 1+0">
                        Глянцевая 1+0
                    </option>
                    <option value="Глянцевая 1+1">
                        Глянцевая 1+1
                    </option>
                </>
            );
        }

        if (id === 'embossing1' || id === 'embossing2') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="фольгой 1 клише">
                        Фольгой 1 клише
                    </option>
                    <option value="фольгой 2 клише">
                        Фольгой 2 клише
                    </option>
                    <option value="конгревное 1 клише">
                        Конгревное 1 клише
                    </option>
                    <option value="конгревное 2 клише">
                        Конгревное 2 клише
                    </option>
                </>
            );
        }

        if (id === 'cutting') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="на формат А3">
                        На формат А3
                    </option>
                    <option value="на формат А4">
                        На формат А4
                    </option>
                </>
            );
        }

        if (id === 'dieCutting') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="1 изделие">1 изделие</option>
                    <option value="2 изделие">2 изделие</option>
                    <option value="3 изделие">3 изделие</option>
                </>
            );
        }

        if (id === 'gluing') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="1 точка">1 точка</option>
                    <option value="2 точки">2 точки</option>
                </>
            );
        }

        if (id === 'binding') {
            return (
                <>
                    <option value="нет">нет</option>
                    <option value="на клей">на клей</option>
                    <option value="на нитку">на нитку</option>
                    <option value="на пружину">на пружину</option>
                </>
            );
        }

        return <option value="нет">нет</option>;
    };

    return (
        <div className="page active">
            <div className="section">
                <h2>Калькуляция тиража</h2>

                <div className="operations">
                    {operations.map((operation) => (
                        <div
                            key={operation.id}
                            className="operation-item"
                        >
                            <label htmlFor={operation.id}>
                                {operation.label}
                            </label>

                            <select
                                id={operation.id}
                                value={
                                    formData[operation.id] || 'нет'
                                }
                                onChange={handleChange}
                            >
                                {getOptions(operation.id)}
                            </select>
                        </div>
                    ))}
                </div>
            </div>

            <div className="section">
                <h2>Дата отгрузки</h2>

                <input
                    type="date"
                    id="shippingDate"
                    value={formData.shippingDate || ''}
                    onChange={handleChange}
                />
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
                    className="next-btn"
                    onClick={onNext}
                    type="button"
                >
                    Далее →
                </button>
            </div>
        </div>
    );
};

export default StepOperations;