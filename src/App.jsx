import { useState } from 'react';
import './styles/style.css';

import StepCustomer from './components/Step/StepCustomer';
import StepMaterials from './components/Step/StepMaterials';
import StepOperations from './components/Step/StepOperations';
import StepOrder from './components/Step/StepOrder';
import StepResults from './components/Step/StepResults';

import { calculateOrder } from './api/googleSheetsApi';

const steps = [
    'customer',
    'materials',
    'operations',
    'order',
    'results',
];

const navigationItems = [
    {
        id: 'customer',
        label: 'Заказчик',
    },
    {
        id: 'materials',
        label: 'Материал',
    },
    {
        id: 'operations',
        label: 'Операции',
    },
    {
        id: 'order',
        label: 'Заказ',
    },
];

const initialFormData = {
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

    shippingDate: '',
};

const initialResults = {
    total: '',
    vat: '',
    final: '',
    sheetsKg: '',
    circulation: '',
    usdRate: '',
    eurRate: '',
};

const requiredFields = [
    ['companyName', 'Введите наименование компании'],
    ['companyAddress', 'Введите адрес'],
    ['companyContacts', 'Введите контакты'],
    ['productName', 'Введите наименование изделия'],
    ['quantity', 'Введите количество изделий'],
    ['perSheet', 'Введите количество изделий на листе'],
    ['materialPrice', 'Введите цену материала за кг'],
    ['printWidth', 'Введите ширину формата печати'],
    ['printHeight', 'Введите высоту формата печати'],
    ['purchaseWidth', 'Введите ширину закупочного формата'],
    ['purchaseHeight', 'Введите высоту закупочного формата'],
    ['shippingDate', 'Выберите дату отгрузки'],
];

const numericFields = [
    ['quantity', 'Количество изделий должно быть больше 0'],
    [
        'perSheet',
        'Количество изделий на листе должно быть больше 0',
    ],
    ['materialPrice', 'Цена материала должна быть больше 0'],
    [
        'printWidth',
        'Ширина формата печати должна быть больше 0',
    ],
    [
        'printHeight',
        'Высота формата печати должна быть больше 0',
    ],
    [
        'purchaseWidth',
        'Ширина закупочного формата должна быть больше 0',
    ],
    [
        'purchaseHeight',
        'Высота закупочного формата должна быть больше 0',
    ],
];

function App() {
    const [step, setStep] = useState('customer');
    const [formData, setFormData] = useState({
        ...initialFormData,
    });
    const [results, setResults] = useState({
        ...initialResults,
    });
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState('');

    const updateFormData = (field, value) => {
        setFormData((prev) => ({
            ...prev,
            [field]: value,
        }));

        setError('');
    };

    const nextStep = () => {
        const currentIndex = steps.indexOf(step);

        if (currentIndex === -1) {
            return;
        }

        if (currentIndex < steps.length - 1) {
            setError('');
            setStep(steps[currentIndex + 1]);
        }
    };

    const prevStep = () => {
        const currentIndex = steps.indexOf(step);

        if (currentIndex === -1) {
            return;
        }

        if (currentIndex > 0) {
            setError('');
            setStep(steps[currentIndex - 1]);
        }
    };

    const validateForm = () => {
        for (const [field, message] of requiredFields) {
            const value = formData[field];

            if (
                value === null ||
                value === undefined ||
                String(value).trim() === ''
            ) {
                return message;
            }
        }

        for (const [field, message] of numericFields) {
            const value = Number(formData[field]);

            if (!Number.isFinite(value) || value <= 0) {
                return message;
            }
        }

        if (
            !formData.printType ||
            formData.printType === 'нет'
        ) {
            return 'Выберите тип печати';
        }

        return '';
    };

    const handleCalculate = async () => {
        const validationError = validateForm();

        if (validationError) {
            setError(validationError);
            return;
        }

        setIsLoading(true);
        setError('');

        try {
            const result = await calculateOrder(formData);

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
                sheetsKg: result.sheetsKg || '',
                circulation: result.circulation || '',
                usdRate: result.usdRate || '',
                eurRate: result.eurRate || '',
            });

            setStep('results');
        } catch (err) {
            console.error('Ошибка расчёта:', err);

            setError(
                err instanceof Error
                    ? err.message
                    : 'Не удалось выполнить расчёт'
            );
        } finally {
            setIsLoading(false);
        }
    };

    const resetApplication = () => {
        setFormData({
            ...initialFormData,
        });

        setResults({
            ...initialResults,
        });

        setError('');
        setIsLoading(false);
        setStep('customer');
    };

    const handleClear = () => {
        const confirmed = window.confirm(
            'Очистить все введённые данные?'
        );

        if (confirmed) {
            resetApplication();
        }
    };

    const handleNewCalculation = () => {
        resetApplication();
    };

    const setNavigationStep = (nextStep) => {
        if (!steps.includes(nextStep)) {
            return;
        }

        setError('');
        setStep(nextStep);
    };

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
                <header>
                    <h1>КАЛЬКУЛЯТОР ПЕЧАТИ</h1>

                    <nav className="navigation">
                        {navigationItems.map((item) => (
                            <button
                                key={item.id}
                                type="button"
                                className={
                                    step === item.id
                                        ? 'nav-btn active'
                                        : 'nav-btn'
                                }
                                onClick={() =>
                                    setNavigationStep(item.id)
                                }
                            >
                                {item.label}
                            </button>
                        ))}
                    </nav>
                </header>

                {error && (
                    <div className="error-message">
                        {error}
                    </div>
                )}

                {renderStep()}
            </main>
        </div>
    );
}

export default App;