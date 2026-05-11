import React from 'react';
import { Navigate, RouteObject } from 'react-router-dom';
import HomePage from '../pages/HomePage';
import NewsPage from '../pages/NewsPage';
import NewsNatureSpotlightPage from '../pages/NewsNatureSpotlightPage';
import NewsJapanTimesAILovePage from '../pages/NewsJapanTimesAILovePage';
import NewsAIAsCatalystWorkshopPage from '../pages/NewsAIAsCatalystWorkshopPage';
import PeoplePage from '../pages/PeoplePage';
import ProjectPage from '../pages/ProjectPage';
import PublicationPage from '../pages/PublicationPage';
import LeaderPage from '../pages/LeaderPage';

export const routes: RouteObject[] = [
  // Canonical English routes
  { path: '/', element: <HomePage /> },
  { path: '/people', element: <PeoplePage /> },
  { path: '/publication', element: <PublicationPage /> },
  { path: '/project', element: <ProjectPage /> },
  { path: '/news', element: <NewsPage /> },
  { path: '/news/ai-as-catalyst-workshop', element: <NewsAIAsCatalystWorkshopPage /> },
  { path: '/news/japantimes-ai-love', element: <NewsJapanTimesAILovePage /> },
  { path: '/news/nature-spotlight', element: <NewsNatureSpotlightPage /> },
  { path: '/leader', element: <LeaderPage locale="en" /> },

  // Canonical Chinese routes
  { path: '/zh', element: <HomePage locale="zh-CN" /> },
  { path: '/zh/people', element: <PeoplePage locale="zh-CN" /> },
  { path: '/zh/publication', element: <PublicationPage locale="zh-CN" /> },
  { path: '/zh/project', element: <ProjectPage locale="zh-CN" /> },
  { path: '/zh/news', element: <NewsPage locale="zh-CN" /> },
  { path: '/zh/news/ai-as-catalyst-workshop', element: <NewsAIAsCatalystWorkshopPage locale="zh-CN" /> },
  { path: '/zh/news/japantimes-ai-love', element: <NewsJapanTimesAILovePage locale="zh-CN" /> },
  { path: '/zh/news/nature-spotlight', element: <NewsNatureSpotlightPage locale="zh-CN" /> },
  { path: '/zh/leader', element: <LeaderPage locale="zh-CN" /> },

  // Compatibility redirects from old .html paths to canonical routes
  { path: '/index.html', element: <Navigate to="/" replace /> },
  { path: '/people.html', element: <Navigate to="/people" replace /> },
  { path: '/publication.html', element: <Navigate to="/publication" replace /> },
  { path: '/legacy/publication.html', element: <Navigate to="/publication" replace /> },
  { path: '/project.html', element: <Navigate to="/project" replace /> },
  { path: '/news.html', element: <Navigate to="/news" replace /> },
  { path: '/leader.html', element: <Navigate to="/leader" replace /> },
  { path: '/news/ai-as-catalyst-workshop.html', element: <Navigate to="/news/ai-as-catalyst-workshop" replace /> },
  { path: '/news/japantimes-ai-love.html', element: <Navigate to="/news/japantimes-ai-love" replace /> },
  { path: '/news/nature-spotlight.html', element: <Navigate to="/news/nature-spotlight" replace /> },
  { path: '/zh/index.html', element: <Navigate to="/zh" replace /> },
  { path: '/zh/people.html', element: <Navigate to="/zh/people" replace /> },
  { path: '/zh/publication.html', element: <Navigate to="/zh/publication" replace /> },
  { path: '/legacy/zh/publication.html', element: <Navigate to="/zh/publication" replace /> },
  { path: '/zh/project.html', element: <Navigate to="/zh/project" replace /> },
  { path: '/zh/news.html', element: <Navigate to="/zh/news" replace /> },
  { path: '/zh/leader.html', element: <Navigate to="/zh/leader" replace /> },
  { path: '/zh/news/ai-as-catalyst-workshop.html', element: <Navigate to="/zh/news/ai-as-catalyst-workshop" replace /> },
  { path: '/zh/news/japantimes-ai-love.html', element: <Navigate to="/zh/news/japantimes-ai-love" replace /> },
  { path: '/zh/news/nature-spotlight.html', element: <Navigate to="/zh/news/nature-spotlight" replace /> },

  { path: '/zh/*', element: <Navigate to="/zh" replace /> },
  { path: '*', element: <Navigate to="/" replace /> }
];
