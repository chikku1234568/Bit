import { Link, Route, Routes } from 'react-router-dom';
import { useEffect, useState } from 'react';
import { getAuthor, setAuthor } from './api';
import { HomePage } from './pages/HomePage';
import { ProjectPage } from './pages/ProjectPage';
import { WhatChangedPage } from './pages/WhatChangedPage';

export default function App() {
  const [author, setAuthorState] = useState(getAuthor());

  useEffect(() => {
    setAuthor(author);
  }, [author]);

  return (
    <div className="app">
      <header className="topbar">
        <Link to="/" className="brand">
          Bit
        </Link>
        <label className="author">
          You are
          <input
            value={author}
            onChange={(e) => setAuthorState(e.target.value)}
            placeholder="demo-user"
          />
        </label>
      </header>
      <main>
        <Routes>
          <Route path="/" element={<HomePage />} />
          <Route path="/projects/:id" element={<ProjectPage />} />
          <Route path="/projects/:id/what-changed" element={<WhatChangedPage />} />
        </Routes>
      </main>
    </div>
  );
}
