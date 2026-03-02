import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';

// Mock window.storage for local development
// (Replaces the Claude artifact storage API with localStorage)
window.storage = {
  _data: {},
  get: async (key) => {
    const val = localStorage.getItem('edp__' + key);
    if (val === null) throw new Error('Key not found');
    return { key, value: val };
  },
  set: async (key, value) => {
    localStorage.setItem('edp__' + key, value);
    return { key, value };
  },
  delete: async (key) => {
    localStorage.removeItem('edp__' + key);
    return { key, deleted: true };
  },
  list: async (prefix = '') => {
    const keys = Object.keys(localStorage)
      .filter(k => k.startsWith('edp__' + prefix))
      .map(k => k.replace('edp__', ''));
    return { keys };
  }
};

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<React.StrictMode><App /></React.StrictMode>);
