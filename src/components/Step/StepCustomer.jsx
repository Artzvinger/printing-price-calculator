import React from 'react';

const StepCustomer = ({ data, onChange, onNext }) => {
    const handleChange = (e) => {
        onChange({ ...data, [e.target.id]: e.target.value });
    };

    return (
        <div className="page active">
            <div className="section">
                <h2>Информация о заказчике</h2>
                <div className="customer-info">
                    <input
                        type="text"
                        id="company-name"
                        placeholder="Наименование компании"
                        value={data.companyName || ''}
                        onChange={handleChange}
                    />
                    <input
                        type="text"
                        id="company-address"
                        placeholder="Адрес"
                        value={data.companyAddress || ''}
                        onChange={handleChange}
                    />
                    <input
                        type="text"
                        id="company-contacts"
                        placeholder="Контакты"
                        value={data.companyContacts || ''}
                        onChange={handleChange}
                    />
                </div>
            </div>
            <div className="page-navigation">
                <button className="next-btn" onClick={onNext}>Далее →</button>
            </div>
        </div>
    );
};

export default StepCustomer;