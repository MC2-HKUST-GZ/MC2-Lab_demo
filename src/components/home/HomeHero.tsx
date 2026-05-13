import React from 'react';
import { Link } from 'react-router-dom';
import { homeText } from '../../data/homePageContent';
import type { Locale } from '../../types/common';
import type { HomePanel } from '../../types/home';
import { runOnEnterOrSpace } from './homeKeyboard';

interface HomeHeroProps {
  locale: Locale;
  onOpenPanel: (panel: Exclude<HomePanel, null>) => void;
}

export default function HomeHero({ locale, onOpenPanel }: HomeHeroProps): JSX.Element {
  const content = homeText[locale];

  return (
    <section className='hero-section'>
      <div
        className='about-text hover-target'
        role='button'
        tabIndex={0}
        onClick={() => onOpenPanel('about')}
        onKeyDown={(event) => runOnEnterOrSpace(event, () => onOpenPanel('about'))}
      >
        About
      </div>
      <div
        className='contact-text hover-target'
        role='button'
        tabIndex={0}
        onClick={() => onOpenPanel('contact')}
        onKeyDown={(event) => runOnEnterOrSpace(event, () => onOpenPanel('contact'))}
      >
        Contact
      </div>
      <div className='section-center'>
        <div className='container-fluid'>
          <div className='row justify-content-center hero-main-row'>
            <div className='col-12 text-center'>
              <h1>Center for Metaverse and Computational Creativity (MC<sup className='hero-brand-sup'>2</sup>)</h1>
            </div>

            <div className='col-12 text-center mt-4 mt-lg-5'>
              <div className='hero-cta-group'>
                <Link className='surreality-btn news-btn hover-target' to='/news' aria-label={content.newsAria}>
                  <i className='fa-regular fa-newspaper btn-icon' aria-hidden='true'></i>
                  <span>{content.newsLabel}</span>
                  <span className='btn-arrow' aria-hidden='true'>→</span>
                </Link>

                <a
                  className='surreality-btn film-btn hover-target'
                  href='https://youtu.be/yCSXbXoK8fg'
                  target='_blank'
                  rel='noopener noreferrer'
                  aria-label='SURREALITY 1.0 Official Exhibition Film'
                >
                  <i className='fa-solid fa-circle-play btn-icon' aria-hidden='true'></i>
                  <span>SURREALITY 1.0 Official Exhibition Film</span>
                  <span className='btn-arrow' aria-hidden='true'>↗</span>
                </a>

                <Link className='surreality-btn open-call-btn hover-target' to='/news?action=showModal' aria-label='SURREALITY 2.0 Open Call'>
                  <i className='fa-solid fa-rocket btn-icon' aria-hidden='true'></i>
                  <span>SURREALITY 2.0 Open Call</span>
                  <span className='btn-arrow' aria-hidden='true'>→</span>
                </Link>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}

