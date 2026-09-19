import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { AddinApp } from './AddinApp';
import './addin.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <AddinApp />
  </StrictMode>,
);
