import React, { useState } from 'react';
import { useLocalStorage } from './hooks/useLocalStorage';
import { calculateOrder } from './api/googleSheetsApi';
import StepCustomer from './components/Step/StepCustomer';
import StepOrder from './components/Step/StepOrder';
import StepMaterials from './components/Step/StepMaterials';
import StepOperations from './components/Step/StepOperations';
import StepResults from './components/Step/StepResults';
import './styles/style.css';

const App = () => {
    const [step, setStep] = useState('customer');
    const [formData, setFormData] = useLocalStorage('calculatorData', {
        companyName: '',
        companyAddress: '',
        companyContacts: '',
        productName: '',
        quantity: '',
        perSheet: '',
        notes: '',
        materialType: 'paper',
        materialPrice: '',
        materialCurrency: '₽',
        formatSize: 'А3',
        printWidth: '',
        printHeight: '',
        purchaseWidth: '',
        purchaseHeight: '',
        usdRate: '',
        eurRate: '',
        opCutting: '',
        opPrinting: '',
        opLamination: '',
        opUV: '',
        opCutting2: '',
        opEmbossing1: '',
        opEmbossing2: '',
        opDieCutting: '',
        opGluing: '',
        opBinding: '',
        shippingDate: '',
    });

    const [results, setResults] = useState(null);
    const [error, setError] = useState(null);
    const [isLoading, setIsLoading] = useState(false);

    const handleChange = (newData) => {
        setFormData(newData);
    };

    const nextStep = () => {
        const steps = ['customer', 'order', 'materials', 'operations', 'results'];
        const currentIndex = steps.indexOf(step);
        if (currentIndex < steps.length - 1) {
            setStep(steps[currentIndex + 1]);
        }
    };

    const prevStep = () => {
        const steps = ['customer', 'order', 'materials', 'operations', 'results'];
        const currentIndex = steps.indexOf(step);
        if (currentIndex > 0) {
            setStep(steps[currentIndex - 1]);
        }
    };

    const handleCalculate = async () => {
        setIsLoading(true);
        try {
            const result = await calculateOrder(formData);
            if (result.success) {
                setResults({
                    total: result.total,
                    vat: result.vat,
                    final: result.final,
                });
                setStep('results');
            } else {
                setError(result.error || 'Ошибка при расчёте');
            }
        } catch (err) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };

    const handleClear = () => {
        if (window.confirm('Очистить все данные?')) {
            setFormData({});
            setResults(null);
            setStep('customer');
        }
    };

    const handleNewCalculation = () => {
        setFormData({});
        setResults(null);
        setStep('customer');
    };

    return (
        <div className="container">
            <header>
                <h1>Калькулятор расчета заказа</h1>
            </header>

            <nav className="navigation">
                {['customer', 'order', 'materials', 'operations', 'results'].map((s) => (
                    <button
                        key={s}
                        className={`nav-btn ${step === s ? 'active' : ''}`}
                        onClick={() => setStep(s)}
                    >
                        {s === 'customer' && 'Заказчик'}
                        {s === 'order' && 'Заказ'}
                        {s === 'materials' && 'Материалы'}
                        {s === 'operations' && 'Операции'}
                        {s === 'results' && 'Результаты'}
                    </button>
                ))}
            </nav>

            {step === 'customer' && (
                <StepCustomer
                    data={formData}
                    onChange={handleChange}
                    onNext={nextStep}
                />
            )}

            {step === 'order' && (
                <StepOrder
                    data={formData}
                    onChange={handleChange}
                    onNext={nextStep}
                    onPrev={prevStep}
                />
            )}

            {step === 'materials' && (
                <StepMaterials
                    data={formData}
                    onChange={handleChange}
                    onNext={nextStep}
                    onPrev={prevStep}
                />
            )}

            {step === 'operations' && (
                <StepOperations
                    data={formData}
                    onChange={handleChange}
                    onCalculate={handleCalculate}
                    onPrev={prevStep}
                    isLoading={isLoading}
                />
            )}

            {step === 'results' && (
                <StepResults
                    results={results}
                    onClear={handleClear}
                    onNew={handleNewCalculation}
                    onPrev={prevStep}
                />
            )}

            {error && (
                <div className="error-message" onClick={() => setError(null)}>
                    ❌ {error}
                </div>
            )}
        </div>
    );
};

export default App;