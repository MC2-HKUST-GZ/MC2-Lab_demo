import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Link, useLocation, useNavigate } from 'react-router-dom';
import { useSectionScrollSpy } from '../../hooks/useSectionScrollSpy';
import type { NewsListContent as NewsListContentType, NewsListItem, NewsModalContent } from '../../types/legacyPages';
import InlineMc2Text from './InlineMc2Text';

interface NewsListContentProps {
  content: NewsListContentType;
}

function NewsCard({ item, onOpenModal }: { item: NewsListItem; onOpenModal: () => void }): JSX.Element {
  const metaSeparatorIndex = (() => {
    const asciiIndex = item.meta.indexOf(':');
    const fullWidthIndex = item.meta.indexOf('：');
    if (asciiIndex === -1) return fullWidthIndex;
    if (fullWidthIndex === -1) return asciiIndex;
    return Math.min(asciiIndex, fullWidthIndex);
  })();
  const metaPrefix = metaSeparatorIndex >= 0 ? item.meta.slice(0, metaSeparatorIndex + 1) : item.meta;
  const metaSuffix = metaSeparatorIndex >= 0 ? item.meta.slice(metaSeparatorIndex + 1).trim() : '';
  const metaJoiner = item.meta.includes('：') ? '' : ' ';
  const cardInner = (
    <div className='row news-row'>
      <div className='col-md-4 news-left'>
        <img
          className='news-cover-photo'
          src={item.image}
          alt={item.imageAlt}
          onError={(event) => { event.currentTarget.src = 'http://www.ivang-design.com/svg-load/portfolio/photo-p.jpg'; }}
        />
      </div>

      <div className='col-md-8'>
        <h4 style={{ color: '#fff', fontSize: item.titleFontSize }}>
          <strong><InlineMc2Text text={item.title} /></strong>
        </h4>
        <p className='news-meta'><strong>{metaPrefix}</strong>{metaSuffix ? `${metaJoiner}${metaSuffix}` : ''}</p>
        {item.descriptions.map((description, index) => (
          <p style={{ color: '#fff', marginBottom: index === item.descriptions.length - 1 && !item.source ? 0 : undefined }} key={description}>
            {description}
          </p>
        ))}
        {item.source ? <p className='news-source'>{item.source}</p> : null}
      </div>
    </div>
  );

  if (item.kind === 'modal') {
    return (
      <button className='button-like' type='button' onClick={onOpenModal}>
        {cardInner}
      </button>
    );
  }

  return (
    <Link className='button-like news-link-card' to={item.href} aria-label={`Open ${item.title}`}>
      {cardInner}
    </Link>
  );
}

function SurrealityModal({ content, isOpen, onClose }: { content: NewsModalContent; isOpen: boolean; onClose: () => void }): JSX.Element {
  if (!isOpen) return <div id='myModal' className='modal'></div>;

  return (
    <div id='myModal' className='modal' style={{ display: 'block' }} onClick={onClose}>
      <div className='modal-content' style={{ overflowY: 'auto', scrollBehavior: 'smooth', maxHeight: '80vh' }} onClick={(event) => event.stopPropagation()}>
        <span className='close' role='button' tabIndex={0} onClick={onClose} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') onClose(); }}>&times;</span>
        <br /><br />
        <img
          className='img-fluid'
          style={{ display: 'block', maxWidth: '100%', height: 'auto' }}
          src={content.poster}
          srcSet={content.srcSet}
          sizes={content.sizes}
          alt={content.posterAlt}
          onError={(event) => {
            event.currentTarget.removeAttribute('srcset');
            event.currentTarget.removeAttribute('sizes');
            event.currentTarget.src = 'http://www.ivang-design.com/svg-load/portfolio/photo-p.jpg';
          }}
        />

        <div className='surreality-cta-wrap'>
          <a className='surreality-cta' href={content.submitHref} target='_blank' rel='noopener noreferrer' aria-label={content.submitAria}>
            <span className='cta-icon' aria-hidden='true'>🚀</span>
            <span className='cta-text'>{content.submitText}</span>
            <span className='cta-arrow' aria-hidden='true'>→</span>
          </a>
          <div className='surreality-cta-sub'>{content.helperText}</div>
        </div>
      </div>
    </div>
  );
}

export default function NewsListContent({ content }: NewsListContentProps): JSX.Element {
  const scrollRef = useRef<HTMLDivElement>(null);
  const navigate = useNavigate();
  const { search } = useLocation();
  const [activeYearId, setActiveYearId] = useState(content.years[0]?.id || '');
  const [isModalOpen, setIsModalOpen] = useState(false);
  const yearIds = useMemo(() => content.years.map((year) => year.id), [content.years]);

  useSectionScrollSpy(scrollRef, yearIds, setActiveYearId);

  useEffect(() => {
    const params = new URLSearchParams(search);
    if (params.get('action') === 'showModal') {
      setIsModalOpen(true);
      navigate(window.location.pathname, { replace: true });
    }
  }, [navigate, search]);

  const selectYear = useCallback((yearId: string) => {
    setActiveYearId(yearId);
    const scrollContainer = scrollRef.current;
    const section = scrollContainer?.querySelector<HTMLElement>(`#${yearId}`);
    if (!scrollContainer || !section) return;
    const containerRect = scrollContainer.getBoundingClientRect();
    const sectionTop = section.getBoundingClientRect().top - containerRect.top + scrollContainer.scrollTop;
    scrollContainer.scrollTo({ top: Math.max(0, sectionTop - 20), behavior: 'smooth' });
  }, []);

  return (
    <>
      <div className='news-section' ref={scrollRef}>
        <div className='container'>
          <div className='row justify-content-center'>
            <div className='col-12 text-center' style={{ paddingTop: '100px', color: '#fff' }}>
              <h3>{content.pageHeading}</h3>
            </div>

            <div className='publication-nav'>
              <div className='container'>
                <ul className='yr__navs'>
                  {content.years.map((year) => (
                    <li className={activeYearId === year.id ? `yr__nav ${year.id} active` : `yr__nav ${year.id}`} key={year.id}>
                      <a href={`#${year.id}`} onClick={(event) => { event.preventDefault(); selectYear(year.id); }}>{year.label.replace('YEAR ', '')}</a>
                    </li>
                  ))}
                </ul>
              </div>
            </div>

            {content.years.map((year) => (
              <section className='news-year-section' id={year.id} key={year.id}>
                <div className='col-12'>
                  <h3 className='news-year-title'>{year.label}</h3>
                </div>
                {year.items.map((item) => (
                  <div className='col-12 mt-3' key={item.id}>
                    <NewsCard item={item} onOpenModal={() => setIsModalOpen(true)} />
                  </div>
                ))}
              </section>
            ))}
          </div>
        </div>
      </div>
      <SurrealityModal content={content.modal} isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} />
    </>
  );
}
