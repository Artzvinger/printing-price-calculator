import React from 'react';

const StepCustomer = ({
                          formData,
                          updateFormData,
                          onNext,
                      }) => {
    const handleChange = (e) => {
        const { id, value } = e.target;

        updateFormData(id, value);
    };

    return (
        <div className="page active">
            <div className="section">
                <h2>Информация о заказчике</h2>

                <div className="customer-info">
                    <input
                        type="text"
                        id="companyName"
                        placeholder="Наименование компании"
                        value={formData.companyName || ''}
                        onChange={handleChange}
                    />

                    <input
                        type="text"
                        id="companyAddress"
                        placeholder="Адрес"
                        value={formData.companyAddress || ''}
                        onChange={handleChange}
                    />

                    <input
                        type="text"
                        id="companyContacts"
                        placeholder="Контакты"
                        value={formData.companyContacts || ''}
                        onChange={handleChange}
                    />
                </div>
            </div>

            <div className="page-navigation">
                <div></div>
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

export default StepCustomer;