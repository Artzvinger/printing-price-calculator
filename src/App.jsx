import { useState } from 'react';
import './styles/style.css';

import StepCustomer from './components/Step/StepCustomer';
import StepMaterials from './components/Step/StepMaterials';
import StepOperations from './components/Step/StepOperations';
import StepOrder from './components/Step/StepOrder';
import StepResults from './components/Step/StepResults';

import { calculateOrder } from './api/googleSheetsApi';

const initialFormData = {
    // Заказчик
    companyName: '',
    companyAddress: '',
    companyContacts: '',

    // Изделие
    productName: '',
    quantity: '',
    perSheet: '',
    notes: '',

    // Материал
    materialType: 'paper',
    materialPrice: '',
    materialCurrency: '₽',

    // Формат
    formatSize: 'А3',
    printWidth: '',
    printHeight: '',
    purchaseWidth: '',
    purchaseHeight: '',

    // Курсы валют
    usdRate: '',
    eurRate: '',

    // Операции
    cuttingFormat: '',
    printType: '',
    lamination: '',
    uvVarnish: '',
    cutting: '',
    embossing1: '',
    embossing2: '',
    dieCutting: '',
    gluing: '',
    binding: '',

    // Срок
    shippingDate: '',
};

const initialResults = {
    total: '',
    vat: '',
    final: '',
};

function App() {
    const [step, setStep] = useState('customer');

    const [formData, setFormData] = useState(initialFormData);

    const [results, setResults] = useState(initialResults);

    const [isLoading, setIsLoading] = useState(false);

    const [error, setError] = useState('');

    // Обновление данных формы
    const updateFormData = (field, value) => {
        setFormData((prev) => ({
            ...prev,
            [field]: value,
        }));

        setError('');
    };

    // Переход вперёд
    const nextStep = () => {
        const steps = [
            'customer',
            'materials',
            'operations',
            'order',
            'results',
        ];

        const currentIndex = steps.indexOf(step);

        if (currentIndex < steps.length - 1) {
            setError('');
            setStep(steps[currentIndex + 1]);
        }
    };

    // Переход назад
    const prevStep = () => {
        const steps = [
            'customer',
            'materials',
            'operations',
            'order',
            'results',
        ];

        const currentIndex = steps.indexOf(step);

        if (currentIndex > 0) {
            setError('');
            setStep(steps[currentIndex - 1]);
        }
    };

    // Расчёт
    const handleCalculate = async () => {
        setIsLoading(true);
        setError('');

        try {
            console.log(
                'Отправляем данные на сервер:',
                formData
            );

            const result = await calculateOrder(formData);

            console.log(
                'Ответ сервера:',
                result
            );

            if (!result || !result.success) {
                throw new Error(
                    result?.error ||
                    result?.message ||
                    'Не удалось выполнить расчёт'
                );
            }

            setResults({
                total: result.total || '0.00 ₽',
                vat: result.vat || '0.00 ₽',
                final: result.final || '0.00 ₽',
            });

            setStep('results');

        } catch (err) {
            console.error(
                'Ошибка расчёта:',
                err
            );

            setError(
                err.message ||
                'Не удалось выполнить расчёт'
            );

        } finally {
            setIsLoading(false);
        }
    };

    // Полная очистка
    const handleClear = () => {
        const confirmed = window.confirm(
            'Очистить все введённые данные?'
        );

        if (!confirmed) {
            return;
        }

        setFormData(initialFormData);
        setResults(initialResults);
        setError('');
        setStep('customer');
    };

    // Новый расчёт
    const handleNewCalculation = () => {
        setFormData(initialFormData);
        setResults(initialResults);
        setError('');
        setStep('customer');
    };

    // Отображение текущего шага
    const renderStep = () => {
        switch (step) {
            case 'customer':
                return (
                    <StepCustomer
                        formData={formData}
                        updateFormData={updateFormData}
                        onNext={nextStep}
                    />
                );

            case 'materials':
                return (
                    <StepMaterials
                        formData={formData}
                        updateFormData={updateFormData}
                        onNext={nextStep}
                        onBack={prevStep}
                    />
                );

            case 'operations':
                return (
                    <StepOperations
                        formData={formData}
                        updateFormData={updateFormData}
                        onNext={nextStep}
                        onBack={prevStep}
                    />
                );

            case 'order':
                return (
                    <StepOrder
                        formData={formData}
                        updateFormData={updateFormData}
                        onCalculate={handleCalculate}
                        onBack={prevStep}
                        isLoading={isLoading}
                    />
                );

            case 'results':
                return (
                    <StepResults
                        formData={formData}
                        results={results}
                        onBack={prevStep}
                        onClear={handleClear}
                        onNew={handleNewCalculation}
                    />
                );

            default:
                return null;
        }
    };

    return (
        <div className="app">

            <main className="container">

                {/* ШАПКА */}

                <header>
                    <h1>КАЛЬКУЛЯТОР ПЕЧАТИ</h1>

                    <nav className="navigation">

                        <button
                            type="button"
                            className={
                                step === 'customer'
                                    ? 'nav-btn active'
                                    : 'nav-btn'
                            }
                            onClick={() => {
                                setError('');
                                setStep('customer');
                            }}
                        >
                            Заказчик
                        </button>

                        <button
                            type="button"
                            className={
                                step === 'materials'
                                    ? 'nav-btn active'
                                    : 'nav-btn'
                            }
                            onClick={() => {
                                setError('');
                                setStep('materials');
                            }}
                        >
                            Материал
                        </button>

                        <button
                            type="button"
                            className={
                                step === 'operations'
                                    ? 'nav-btn active'
                                    : 'nav-btn'
                            }
                            onClick={() => {
                                setError('');
                                setStep('operations');
                            }}
                        >
                            Операции
                        </button>

                        <button
                            type="button"
                            className={
                                step === 'order'
                                    ? 'nav-btn active'
                                    : 'nav-btn'
                            }
                            onClick={() => {
                                setError('');
                                setStep('order');
                            }}
                        >
                            Заказ
                        </button>

                    </nav>
                </header>

                {/* ОШИБКА */}

                {error && (
                    <div className="error-message">
                        {error}
                    </div>
                )}

                {/* ТЕКУЩИЙ ШАГ */}

                {renderStep()}

            </main>

        </div>
    );
}

export default App;