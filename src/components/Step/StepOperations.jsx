import React from 'react';

const operations = [
    { id: 'material-type', label: 'Материал:' },
    { id: 'cutting-format', label: 'Подрезка:' },
    { id: 'print-type', label: 'Печать:' },
    { id: 'lamination', label: 'Ламинация:' },
    { id: 'uv-varnish', label: 'УФ-лак:' },
    { id: 'cutting', label: 'Резка:' },
    { id: 'embossing1', label: 'Тиснение (1):' },
    { id: 'embossing2', label: 'Тиснение (2):' },
    { id: 'die-cutting', label: 'Вырубка:' },
    { id: 'gluing', label: 'Склейка:' },
    { id: 'binding', label: 'Брошюровка:' },
];

const StepOperations = ({ data, onChange, onCalculate, onPrev, isLoading }) => {
    const handleChange = (e) => {
        onChange({ ...data, [e.target.id]: e.target.value });
    };

    const handleSubmit = (e) => {
        e.preventDefault();
        onCalculate();
    };

    return (
        <div className="page active">
            <form onSubmit={handleSubmit}>
                <div className="section">
                    <h2>Калькуляция тиража</h2>
                    <div className="operations">
                        {operations.map((op) => (
                            <div key={op.id} className="operation-item">
                                <label htmlFor={op.id}>{op.label}</label>
                                <select
                                    id={op.id}
                                    value={data[op.id] || ''}
                                    onChange={handleChange}
                                >
                                    <option value="нет">нет</option>
                                    {op.id === 'material-type' && (
                                        <>
                                            <option value="Бумага_мелованная 150 г/м²">Бумага_мелованная 150 г/м²</option>
                                            <option value="Картон 220 г/м²">Картон 220 г/м²</option>
                                        </>
                                    )}
                                    {op.id === 'cutting-format' && (
                                        <>
                                            <option value="на формат А2">На формат А2</option>
                                            <option value="на формат А3">На формат А3</option>
                                            <option value="на формат А4">На формат А4</option>
                                        </>
                                    )}
                                    {op.id === 'print-type' && (
                                        <>
                                            <option value="1+0">1+0</option>
                                            <option value="2+0">2+0</option>
                                            <option value="3+0">3+0</option>
                                            <option value="4+0">4+0</option>
                                            <option value="5+0">5+0</option>
                                        </>
                                    )}
                                    {['lamination', 'uv-varnish'].includes(op.id) && (
                                        <>
                                            <option value="Глянцевая 1+0">Глянцевая 1+0</option>
                                            <option value="Глянцевая 1+1">Глянцевая 1+1</option>
                                        </>
                                    )}
                                    {['embossing1', 'embossing2'].includes(op.id) && (
                                        <>
                                            <option value="фольгой 1 клише">Фольгой 1 клише</option>
                                            <option value="фольгой 2 клише">Фольгой 2 клише</option>
                                            <option value="конгревное 1 клише">Конгревное 1 клише</option>
                                            <option value="конгревное 2 клише">Конгревное 2 клише</option>
                                        </>
                                    )}
                                    {op.id === 'cutting' && (
                                        <>
                                            <option value="на формат А3">На формат А3</option>
                                            <option value="на формат А4">На формат А4</option>
                                        </>
                                    )}
                                    {op.id === 'die-cutting' && (
                                        <>
                                            <option value="1 изделие">1 изделие</option>
                                            <option value="2 изделие">2 изделие</option>
                                            <option value="3 изделие">3 изделие</option>
                                        </>
                                    )}
                                    {op.id === 'gluing' && (
                                        <>
                                            <option value="1 точка">1 точка</option>
                                            <option value="2 точки">2 точки</option>
                                        </>
                                    )}
                                    {op.id === 'binding' && (
                                        <>
                                            <option value="на клей">на клей</option>
                                            <option value="на нитку">на нитку</option>
                                            <option value="на пружину">на пружину</option>
                                        </>
                                    )}
                                </select>
                            </div>
                        ))}
                    </div>
                </div>

                <div className="section">
                    <h2>Дата отгрузки</h2>
                    <input
                        type="date"
                        id="shipping-date"
                        value={data.shippingDate || ''}
                        onChange={handleChange}
                    />
                </div>

                <div className="page-navigation">
                    <button className="prev-btn" onClick={onPrev} type="button">← Назад</button>
                    <button className="next-btn" type="submit" disabled={isLoading}>
                        {isLoading ? '⏳ Расчет...' : 'Рассчитать →'}
                    </button>
                </div>
            </form>
        </div>
    );
};

export default StepOperations;