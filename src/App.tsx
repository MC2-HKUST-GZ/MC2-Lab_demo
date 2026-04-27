import React from 'react';
import { useRoutes } from 'react-router-dom';
import { routes } from './router/routes';

export default function App(): JSX.Element {
  return <>{useRoutes(routes)}</>;
}
