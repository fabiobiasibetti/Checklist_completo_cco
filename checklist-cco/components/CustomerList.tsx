import React, { useState } from 'react';
import { Customer } from '../types';
import { saveCustomers } from '../services/storageService';
import { Plus, Trash2, Building2, Phone, Mail, User } from 'lucide-react';

interface CustomerListProps {
  customers: Customer[];
  setCustomers: React.Dispatch<React.SetStateAction<Customer[]>>;
}

const CustomerList: React.FC<CustomerListProps> = ({ customers, setCustomers }) => {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [formData, setFormData] = useState<Partial<Customer>>({
    name: '',
    company: '',
    email: '',
    phone: '',
    status: 'Lead'
  });

  const handleSave = () => {
    if (!formData.name) return;
    const newCustomer: Customer = {
        ...formData as Customer,
        id: Date.now().toString(),
    };
    const updated = [...customers, newCustomer];
    setCustomers(updated);
    saveCustomers(updated);
    setIsFormOpen(false);
    setFormData({ name: '', company: '', email: '', phone: '', status: 'Lead' });
  };

  const handleDelete = (id: string) => {
      if (confirm('Excluir este cliente?')) {
          const updated = customers.filter(c => c.id !== id);
          setCustomers(updated);
          saveCustomers(updated);
      }
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center bg-white p-4 rounded-xl shadow-sm border border-gray-100">
        <h2 className="text-xl font-bold text-gray-800">Clientes & Leads</h2>
        <button 
          onClick={() => setIsFormOpen(true)}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition-colors"
        >
          <Plus size={18} />
          <span>Novo Cliente</span>
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {customers.map(customer => (
          <div key={customer.id} className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 hover:shadow-md transition-shadow relative group">
            <button 
                onClick={() => handleDelete(customer.id)}
                className="absolute top-4 right-4 text-gray-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity"
            >
                <Trash2 size={18} />
            </button>
            <div className="flex items-center gap-3 mb-4">
                <div className="w-10 h-10 rounded-full bg-blue-100 flex items-center justify-center text-blue-600 font-bold">
                    {customer.name.charAt(0)}
                </div>
                <div>
                    <h3 className="font-bold text-gray-800">{customer.name}</h3>
                    <span className={`text-xs px-2 py-0.5 rounded-full ${customer.status === 'Ativo' ? 'bg-green-100 text-green-700' : 'bg-gray-100 text-gray-600'}`}>
                        {customer.status}
                    </span>
                </div>
            </div>
            
            <div className="space-y-2 text-sm text-gray-600">
                <div className="flex items-center gap-2">
                    <Building2 size={16} className="text-gray-400"/>
                    <span>{customer.company || '-'}</span>
                </div>
                <div className="flex items-center gap-2">
                    <Mail size={16} className="text-gray-400"/>
                    <a href={`mailto:${customer.email}`} className="hover:text-blue-600">{customer.email || '-'}</a>
                </div>
                <div className="flex items-center gap-2">
                    <Phone size={16} className="text-gray-400"/>
                    <span>{customer.phone || '-'}</span>
                </div>
            </div>
          </div>
        ))}
      </div>

       {/* Simple Customer Form Modal */}
       {isFormOpen && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-md animate-in fade-in zoom-in duration-200">
            <div className="p-6 border-b border-gray-100">
                <h3 className="text-lg font-bold text-gray-800">Novo Cliente</h3>
            </div>
            <div className="p-6 space-y-4">
                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Nome</label>
                    <input type="text" value={formData.name} onChange={e => setFormData({...formData, name: e.target.value})} className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Empresa</label>
                    <input type="text" value={formData.company} onChange={e => setFormData({...formData, company: e.target.value})} className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Email</label>
                        <input type="email" value={formData.email} onChange={e => setFormData({...formData, email: e.target.value})} className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </div>
                     <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">Telefone</label>
                        <input type="text" value={formData.phone} onChange={e => setFormData({...formData, phone: e.target.value})} className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" />
                    </div>
                </div>
                <div>
                     <label className="block text-sm font-medium text-gray-700 mb-1">Status</label>
                     <select value={formData.status} onChange={e => setFormData({...formData, status: e.target.value as any})} className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <option>Lead</option>
                        <option>Ativo</option>
                        <option>Inativo</option>
                     </select>
                </div>
            </div>
            <div className="p-6 border-t border-gray-100 flex justify-end gap-3">
                <button onClick={() => setIsFormOpen(false)} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded-lg font-medium transition-colors">Cancelar</button>
                <button onClick={handleSave} className="px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors">Salvar Cliente</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default CustomerList;
