import React, { useEffect, useRef } from 'react';
import SitePageShell from '../components/SitePageShell';
import type { Locale } from '../types/common';
import '../styles/workshop-page.css';

interface NewsAIAsCatalystWorkshopPageProps {
  locale?: Locale;
}

const FALLBACK_PHOTO = 'http://www.ivang-design.com/svg-load/portfolio/photo-p.jpg';

function useWorkshopSpeakerCards(pageRef: React.RefObject<HTMLElement>): void {
  useEffect(() => {
    const page = pageRef.current;
    if (!page) return undefined;

    const speakerCards = Array.from(page.querySelectorAll<HTMLElement>('.speaker-card-news'));
    if (!speakerCards.length) return undefined;

    const collapsedHintLabels = new Map<HTMLElement, string>();
    const expandedHintLabels = new Map<HTMLElement, string>();

    speakerCards.forEach((card) => {
      const hint = card.querySelector<HTMLElement>('.speaker-card-news__hint');
      const collapsedLabel = hint?.textContent?.trim() || 'Click to view bio';
      collapsedHintLabels.set(card, collapsedLabel);
      expandedHintLabels.set(card, collapsedLabel === '点击查看简介' ? '收起简介' : 'Hide bio');
    });

    const setCardState = (card: HTMLElement, shouldOpen: boolean) => {
      const details = card.querySelector<HTMLElement>('.speaker-card-news__details');
      const hint = card.querySelector<HTMLElement>('.speaker-card-news__hint');

      card.classList.toggle('is-open', shouldOpen);
      card.setAttribute('aria-expanded', shouldOpen ? 'true' : 'false');

      if (hint) {
        hint.textContent = shouldOpen
          ? expandedHintLabels.get(card) || 'Hide bio'
          : collapsedHintLabels.get(card) || 'Click to view bio';
      }

      if (details) {
        details.style.maxHeight = shouldOpen ? details.scrollHeight + 'px' : '0px';
      }
    };

    const scrollCardIntoView = (card: HTMLElement) => {
      const scrollContainer = page.querySelector<HTMLElement>('.news-section');
      if (!scrollContainer) return;

      const topOffset = 108;
      const containerTop = scrollContainer.getBoundingClientRect().top;
      const cardTop = card.getBoundingClientRect().top - containerTop + scrollContainer.scrollTop - topOffset;
      scrollContainer.scrollTo({
        top: Math.max(0, cardTop),
        behavior: 'smooth'
      });
    };

    const clickHandlers = new Map<HTMLElement, EventListener>();
    const keydownHandlers = new Map<HTMLElement, EventListener>();

    speakerCards.forEach((card) => {
      setCardState(card, false);

      const clickHandler = () => {
        const shouldOpen = !card.classList.contains('is-open');
        speakerCards.forEach((otherCard) => {
          setCardState(otherCard, otherCard === card ? shouldOpen : false);
        });
      };

      const keydownHandler = (event: Event) => {
        const keyboardEvent = event as KeyboardEvent;
        if (keyboardEvent.key === 'Enter' || keyboardEvent.key === ' ') {
          keyboardEvent.preventDefault();
          card.click();
        }
      };

      clickHandlers.set(card, clickHandler);
      keydownHandlers.set(card, keydownHandler);
      card.addEventListener('click', clickHandler);
      card.addEventListener('keydown', keydownHandler);
    });

    const openCardByHash = () => {
      if (!window.location.hash) return;

      const targetCard = page.querySelector<HTMLElement>(window.location.hash + '.speaker-card-news');
      if (!targetCard) return;

      speakerCards.forEach((card) => {
        setCardState(card, card === targetCard);
      });

      requestAnimationFrame(() => {
        targetCard.focus({ preventScroll: true });
        scrollCardIntoView(targetCard);

        window.setTimeout(() => {
          scrollCardIntoView(targetCard);
        }, 240);
      });
    };

    const resizeHandler = () => {
      speakerCards.forEach((card) => {
        if (!card.classList.contains('is-open')) return;
        const details = card.querySelector<HTMLElement>('.speaker-card-news__details');
        if (details) {
          details.style.maxHeight = details.scrollHeight + 'px';
        }
      });
    };

    window.addEventListener('resize', resizeHandler);
    window.addEventListener('hashchange', openCardByHash);
    openCardByHash();

    return () => {
      window.removeEventListener('resize', resizeHandler);
      window.removeEventListener('hashchange', openCardByHash);
      speakerCards.forEach((card) => {
        const clickHandler = clickHandlers.get(card);
        const keydownHandler = keydownHandlers.get(card);
        if (clickHandler) card.removeEventListener('click', clickHandler);
        if (keydownHandler) card.removeEventListener('keydown', keydownHandler);
      });
    };
  }, [pageRef]);
}

function hideOrganizerBrandImage(event: React.SyntheticEvent<HTMLImageElement>): void {
  const image = event.currentTarget;
  image.style.display = 'none';
  image.closest('.organizer-brand')?.classList.add('organizer-brand--text-only');
}

function hideAffiliationTileImage(event: React.SyntheticEvent<HTMLImageElement>): void {
  const image = event.currentTarget;
  image.style.display = 'none';
  image.closest('.affiliation-tile')?.classList.add('affiliation-tile--text-only');
}

function useFallbackPhoto(event: React.SyntheticEvent<HTMLImageElement>): void {
  const image = event.currentTarget;
  if (image.src === FALLBACK_PHOTO) return;
  image.src = FALLBACK_PHOTO;
}

function EnglishWorkshopArticle(): JSX.Element {
  return (
    <div className='news-section'>
      <div className='container'>
        <div className='row justify-content-center'>
          <div className='col-12 text-center' style={{ paddingTop: '100px', color: '#fff' }}>
            <h3>{"News"}</h3>
          </div>
          <div className='col-lg-10 mt-4 mb-5'>
            <article className='news-article-card'>
              <h2 className='article-title'>{"AI as Catalyst Workshop to Explore AI in Healthcare and Education at HKUST(GZ)"}</h2>
              <p className='article-meta'>
                <strong>{"Published:"}</strong>
                {" April 10, 2026"}
              </p>
              <section className='affiliation-showcase' aria-label='Workshop organizers and participating institutions'>
                <div className='affiliation-group'>
                  <p className='affiliation-group__title'>{"Organizers"}</p>
                  <div className='organizer-lockup'>
                    <div className='organizer-lockup__row organizer-lockup__row--single'>
                      <div className='organizer-pair organizer-pair--schools'>
                        <div className='organizer-brand organizer-brand--mit'>
                          <div className='organizer-brand__visual'>
                            <img src='https://1000logos.net/wp-content/uploads/2026/01/MIT-logo.png' alt='MIT logo' className='organizer-brand__logo organizer-brand__logo--mit' />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"Massachusetts Institute of Technology"}</span>
                            <span className='organizer-brand__en'>{"MIT"}</span>
                          </div>
                        </div>
                        <span className='organizer-pair__mark' aria-hidden='true'>{"×"}</span>
                        <div className='organizer-brand organizer-brand--hkustgz'>
                          <div className='organizer-brand__visual'>
                            <img src='/images/UST-GZ-EN.png' alt='HKUST(GZ) logo' className='organizer-brand__logo organizer-brand__logo--hkustgz' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"The Hong Kong University of Science and Technology (Guangzhou)"}</span>
                            <span className='organizer-brand__en'>{"HKUST(GZ)"}</span>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className='organizer-lockup__row organizer-lockup__row--single'>
                      <div className='organizer-pair organizer-pair--institutions'>
                        <div className='organizer-brand organizer-brand--critical-data'>
                          <div className='organizer-brand__visual'>
                            <img src='https://criticaldatathon.github.io/static/img/logo.svg' alt='MIT Critical Data logo' className='organizer-brand__logo organizer-brand__logo--critical-data' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"MIT Critical Data"}</span>
                          </div>
                        </div>
                        <span className='organizer-pair__mark' aria-hidden='true'>{"×"}</span>
                        <div className='organizer-brand organizer-brand--cma'>
                          <div className='organizer-brand__visual'>
                            <img src='/images/workshop-logos/cma-official.png' alt='Computational Media and Arts Thrust logo' className='organizer-brand__logo organizer-brand__logo--cma' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"Computational Media and Arts"}</span>
                          </div>
                        </div>
                        <span className='organizer-pair__mark' aria-hidden='true'>{"×"}</span>
                        <div className='organizer-brand organizer-brand--mc2'>
                          <div className='organizer-brand__visual'>
                            <img src='/images/MC2.png' alt='MC2 logo' className='organizer-brand__logo organizer-brand__logo--mc2' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>
                              {"MC"}
                              <sup>{"2"}</sup>
                              {" Lab"}
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
                <div className='affiliation-group affiliation-group--is-hidden'>
                  <p className='affiliation-group__title'>{"Participating Institutions"}</p>
                  <div className='affiliation-grid affiliation-grid--participants'>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--dark'>
                        <img src='/images/UST-EN.png' alt='HKUST main campus logo' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"HKUST"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/harvard.png' alt='Harvard University logo' className='affiliation-tile__logo affiliation-tile__logo--participant affiliation-tile__logo--participant-crest affiliation-tile__logo--participant-harvard' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"Harvard"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/nus.png' alt='National University of Singapore logo' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"National University of Singapore"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/hku.png' alt='The University of Hong Kong logo' className='affiliation-tile__logo affiliation-tile__logo--participant affiliation-tile__logo--participant-crest' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"The University of Hong Kong"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/cuhk.png' alt='The Chinese University of Hong Kong logo' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"The Chinese University of Hong Kong"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/emory.png' alt='Emory University logo' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"Emory University"}</p>
                    </article>
                  </div>
                </div>
              </section>
              <img className='article-cover' src='/images/20260410_ai_as_catalyst_workshop_cover.jpg' alt='HKUST(GZ) campus image for the AI as Catalyst workshop.' onError={useFallbackPhoto} />
              <div className='article-info-grid'>
                <div className='article-info-box'>
                  <span className='info-label'>{"Event Date"}</span>
                  <span className='info-value'>{"May 12, 2026"}</span>
                </div>
                <div className='article-info-box'>
                  <span className='info-label'>{"Location"}</span>
                  <span className='info-value'>{"Guangzhou, China"}</span>
                </div>
                <div className='article-info-box'>
                  <span className='info-label'>{"Format"}</span>
                  <span className='info-value'>{"Full-day workshop and panel"}</span>
                </div>
              </div>
              <p>
                <strong>{"AI as Catalyst"}</strong>
                {" is a full-day workshop taking place on May 12, 2026 at HKUST(GZ), bringing together students, educators, clinicians, and researchers to examine how artificial intelligence should be designed, evaluated, and governed in healthcare and education.\n                    "}
              </p>
              <p>{"\n                        The program combines hands-on experimentation, interdisciplinary discussion, and collaborative design. Rather than treating AI as a shortcut for efficiency, the workshop is framed around critical engagement: participants will test capabilities, examine failure modes, and discuss the social and institutional safeguards needed for responsible deployment.\n                    "}</p>
              <h4>{"Program Snapshot"}</h4>
              <div className='agenda-snapshot'>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"09:00-09:30"}</span>
                  <div>
                    <strong>{"Registration & Welcome"}</strong>
                    <p>{"Check-in."}</p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"09:30-09:45"}</span>
                  <div>
                    <strong>{"Opening Ceremony"}</strong>
                    <p>
                      {"Workshops Overview / Opening Remarks. Leads: "}
                      <a href='#speaker-pan-hui' className='agenda-speaker-link'>{"Pan Hui"}</a>
                      {" and "}
                      <a href='#speaker-leo-anthony-celi' className='agenda-speaker-link'>{"Leo Anthony Celi"}</a>
                      {"."}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"09:45-11:05"}</span>
                  <div>
                    <strong>{"Workshop 1"}</strong>
                    <p>
                      {"The Art of Healing: Creative Arts as Pedagogical Practice. Lead: "}
                      <a href='#speaker-tengjia-zuo' className='agenda-speaker-link'>{"Tengjia Zuo"}</a>
                      {"."}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"11:05-11:15"}</span>
                  <div>
                    <strong>{"Break"}</strong>
                    <p>{"Networking."}</p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"11:15-12:45"}</span>
                  <div>
                    <strong>{"Workshop 2"}</strong>
                    <p>
                      {"Health AI Systems Thinking for Community (HASTC). Lead: "}
                      <a href='#speaker-mornin-feng' className='agenda-speaker-link'>{"Mornin Feng"}</a>
                      {"."}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"12:45-14:35"}</span>
                  <div>
                    <strong>
                      {"Lunch & "}
                      <a href='https://surreality.world/' target='_blank' rel='noopener noreferrer'>
                        <em>{"Surreality Exhibition"}</em>
                      </a>
                      {" Viewing"}
                    </strong>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"14:35-15:35"}</span>
                  <div>
                    <strong>{"Panel Discussion"}</strong>
                    <p>
                      {"Panel Title: From Personal Agency to Systemic Safety: Real-World AI in Healthcare. Lead: "}
                      <a href='#speaker-qiushi-zhou' className='agenda-speaker-link'>{"Qiushi Zhou"}</a>
                      {"."}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"15:35-15:55"}</span>
                  <div>
                    <strong>
                      <a href='#sponsor-wsc' className='agenda-speaker-link'>{"WSC Coffee Break"}</a>
                    </strong>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"15:55-17:15"}</span>
                  <div>
                    <strong>{"Workshop 3"}</strong>
                    <p>
                      {"LLM-a-thon: From Use Case to Ground Truth. Lead: "}
                      <a href='#speaker-calvin-kalun-or' className='agenda-speaker-link'>{"Calvin K.L. Or"}</a>
                      {"."}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"17:15-17:30"}</span>
                  <div>
                    <strong>{"Closing"}</strong>
                    <p>
                      {"Key Takeaways / Next Steps / Wrap-Up. Leads: "}
                      <a href='#speaker-pan-hui' className='agenda-speaker-link'>{"Pan Hui"}</a>
                      {" and "}
                      <a href='#speaker-leo-anthony-celi' className='agenda-speaker-link'>{"Leo Anthony Celi"}</a>
                      {"."}
                    </p>
                  </div>
                </div>
              </div>
              <h4>{"Speakers and Organizers"}</h4>
              <p>{"\n                        The workshop features an interdisciplinary group of speakers and organizers from HKUST(GZ), HKUST, MIT, Harvard, the National University of Singapore, The University of Hong Kong, The Chinese University of Hong Kong, Emory University, and partner institutions. Click any card to expand the short biography.\n                    "}</p>
              <div className='speaker-grid-news'>
                <article className='speaker-card-news' id='speaker-pan-hui' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/pan_hui.jpg?v=20260412-2' alt='Pan Hui' className='speaker-card-news__image speaker-card-news__image--pan-hui' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Pan Hui"}</h5>
                    <p className='speaker-card-news__role'>{"Chair Professor"}</p>
                    <p className='speaker-card-news__org'>
                      {"HKUST(GZ) / MC"}
                      <sup>{"2"}</sup>
                      {" Lab"}
                    </p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"FREng (International Fellow of the Royal Academy of Engineering), IEEE Fellow, ACM Distinguished Scientist, Member of Academia Europaea, and Chair Professor of Computational Media and Arts at HKUST(GZ). His research focuses on ubiquitous computing, mobile computing, augmented/virtual reality, data science, social and mobile networks, and metaverse-related systems."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-leo-anthony-celi' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/leo_celi.jpg' alt='Leo Anthony Celi' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Leo Anthony Celi"}</h5>
                    <p className='speaker-card-news__role'>{"Senior Research Scientist & Associate Professor"}</p>
                    <p className='speaker-card-news__org'>{"MIT / Harvard"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Senior Research Scientist at Massachusetts Institute of Technology, Clinical Research Director of the MIT Laboratory for Computational Physiology, Co-Director of MIT Sana, Staff Physician in the Division of Pulmonary, Critical Care and Sleep Medicine at Beth Israel Deaconess Medical Center, and Part-time Associate Professor of Medicine at Harvard Medical School. His work focuses on using data science and AI to improve critical care, with particular emphasis on open clinical data and reducing bias in healthcare systems."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/Chenlei.png' alt='Lei Chen' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Lei Chen"}</h5>
                    <p className='speaker-card-news__role'>{"Dean / Chair Professor"}</p>
                    <p className='speaker-card-news__org'>{"HKUST(GZ) / Information Hub"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"ACM Fellow, IEEE Fellow, Dean of Information Hub at The Hong Kong University of Science and Technology (Guangzhou), Director of HKUST Big Data Institute, and Chair Professor in the Thrust of Data Science and Analytics and the Thrust of Artificial Intelligence at HKUST (Guangzhou). His research interests include data-driven machine learning, crowd-sourcing-based data processing, uncertain and probabilistic databases, web information management, multimedia systems, knowledge graphs, blockchain, data privacy, and spatial-temporal data management."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-mornin-feng' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/mornin_feng.jpg' alt='Mornin Feng' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Mornin Feng"}</h5>
                    <p className='speaker-card-news__role'>{"Associate Professor"}</p>
                    <p className='speaker-card-news__org'>{"National University of Singapore"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Associate Professor at the National University of Singapore (Saw Swee Hock School of Public Health), Director of the AI for Public Health (AI4PH) Programme, Domain Lead for Biostatistics, Modelling, AI and Data Analytics (B.MAD), Associate Professor (Joint) at the Yong Loo Lin School of Medicine, Biomedical Engineering Department, and Institute of Data Science, and Chair of the Singapore Chapter of Observational Health Data Sciences and Informatics (OHDSI). His research focuses on healthcare AI, machine learning for clinical data, and data-driven decision-making, with applications in medical imaging, treatment optimization, and clinical NLP."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-calvin-kalun-or' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/calvin_or.jpg' alt='Calvin Kalun Or' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Calvin Kalun Or"}</h5>
                    <p className='speaker-card-news__role'>{"Associate Professor"}</p>
                    <p className='speaker-card-news__org'>{"The University of Hong Kong"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Associate Professor and Assistant Head of Department at the The University of Hong Kong (Department of Data and Systems Engineering). His research focuses on human factors and ergonomics, human-computer interaction, and healthcare informatics, particularly the design, implementation, and evaluation of health information technologies to improve patient safety, system performance, and quality of care."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/Lowell_Ling.jpg' alt='Lowell Ling' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Lowell Ling"}</h5>
                    <p className='speaker-card-news__role'>{"Associate Professor"}</p>
                    <p className='speaker-card-news__org'>{"The Chinese University of Hong Kong"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Associate Professor in the Department of Anaesthesia and Intensive Care at the The Chinese University of Hong Kong, and Assistant Professor in the Intensive Care Unit at Prince of Wales Hospital. His research focuses on sepsis, organ dysfunction, and critical care, particularly the epidemiology of sepsis in Hong Kong and its genomic mechanisms."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/xiangliangzhang.jpg' alt='Nancy Zhang' className='speaker-card-news__image speaker-card-news__image--xianglilan' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Nancy Zhang"}</h5>
                    <p className='speaker-card-news__role'>{"Assistant Professor"}</p>
                    <p className='speaker-card-news__org'>{"The Chinese University of Hong Kong"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Assistant Professor at the Faculty of Medicine, The Chinese University of Hong Kong. Her research focuses on applying artificial intelligence and machine learning to clinical diagnosis, disease surveillance, and treatment decision support, with applications in precision medicine and critical care."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-qiushi-zhou' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/qiushi_zhou_20260423.png' alt='Qiushi Zhou' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Qiushi Zhou"}</h5>
                    <p className='speaker-card-news__role'>{"Assistant Professor"}</p>
                    <p className='speaker-card-news__org'>{"HKUST(GZ)"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Assistant Professor at the Hong Kong University of Science and Technology (Guangzhou), affiliated with the Thrust of Computational Media and Arts and the Thrust of Internet of Things. His research focuses on Human-Computer Interaction and Extended Reality, particularly novel interaction techniques integrating XR, AI, and IoT."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/xin_tong.jpg' alt='Xin Tong' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Xin Tong"}</h5>
                    <p className='speaker-card-news__role'>{"Assistant Professor"}</p>
                    <p className='speaker-card-news__org'>{"HKUST(GZ)"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Assistant Professor at Thrusts of Computational Media and Arts (CMA), Hong Kong University of Science and Technology (Guangzhou). She is working on human-computer interaction and human-AI collaboration, with a focus on healthcare, wellbeing, and accessibility using technologies such as VR/AR and generative AI."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-tengjia-zuo' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/tengjia_zuo.jpg' alt='Tengjia Zuo' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Tengjia Zuo"}</h5>
                    <p className='speaker-card-news__role'>{"Assistant Professor"}</p>
                    <p className='speaker-card-news__org'>{"HKUST(GZ)"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Assistant Professor at Computational Media and Arts, Hong Kong University of Science and Technology (Guangzhou). Her research focuses on mixed reality, serious games, and human-computer interaction, particularly game-based learning and player experience design."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/Hyunjung.jpg' alt='Hyunjung Gloria Kwak' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Hyunjung Gloria Kwak"}</h5>
                    <p className='speaker-card-news__role'>{"Assistant Professor"}</p>
                    <p className='speaker-card-news__org'>{"Emory University"}</p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Assistant Professor at Emory University's School of Nursing. Her research focuses on applying artificial intelligence and machine learning to improve clinical decision-making, with an emphasis on advancing equity and reducing health disparities in healthcare."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/james_formal.jpg' alt='James Yiming Zhu' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"James Yiming Zhu"}</h5>
                    <p className='speaker-card-news__role'>{"PhD Student"}</p>
                    <p className='speaker-card-news__org'>
                      {"HKUST(GZ) / MC"}
                      <sup>{"2"}</sup>
                      {" Lab"}
                    </p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Ph.D. candidate at the Hong Kong University of Science and Technology, affiliated with the MC² Lab. His research interests lie in social computing, data science, and natural language processing."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/yulin_yao.jpg' alt='Yulin Yao' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Yulin Yao"}</h5>
                    <p className='speaker-card-news__role'>{"PhD Student / Visual Artist"}</p>
                    <p className='speaker-card-news__org'>
                      {"HKUST(GZ) / MC"}
                      <sup>{"2"}</sup>
                      {" Lab"}
                    </p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Yulin Yao is a PhD student in Computational Media and Arts at The Hong Kong University of Science and Technology (Guangzhou) and a digital visual artist. Her work operates at the intersection of immersive storytelling, VR and HCI. Grounded in spatial design and exhibition practice across Europe and China, she creates immersive digital environments exploring technology's impact on memory, emotion, and culture. Her practice bridges computational media and visual art, treating space itself as a medium for psychological inquiry and immersive world-building."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/bianca_yang.jpg' alt='Bianca Ruoshan Yang' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Bianca Ruoshan Yang"}</h5>
                    <p className='speaker-card-news__role'>{"PhD Student / Musician"}</p>
                    <p className='speaker-card-news__org'>
                      {"HKUST(GZ) / MC"}
                      <sup>{"2"}</sup>
                      {" Lab"}
                    </p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Bianca Ruoshan Yang is a Ph.D. Student at Hong Kong University of Science and Technology, affiliated with the MC2 Lab. Her research sits at the intersection of Multimodal AI, Learning-Centered XR Design, and Intangible Cultural Heritage."}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/primo_formal.jpeg' alt='Primo Dongyijie Pan' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Primo Dongyijie Pan"}</h5>
                    <p className='speaker-card-news__role'>{"MPhil Student"}</p>
                    <p className='speaker-card-news__org'>
                      {"HKUST(GZ) / MC"}
                      <sup>{"2"}</sup>
                      {" Lab"}
                    </p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>
                        {"Primo's research explores how established psychotherapeutic approaches, especially cognitive behavioral therapy, can be translated into the language of human-computer interaction. He has contributed to projects including Lingolift and Virtual AI-SP, with work accepted at venues such as CHI and recognized by the China Computer Federation and the Tencent Light Foundation. In his future doctoral research, he aims to combine AI and XR with emerging lifestyle intervention methods to create new possibilities for chronic disease management. Personal homepage: "}
                        <a href='https://primopan.github.io/about/' target='_blank' rel='noopener noreferrer'>{"https://primopan.github.io/about/"}</a>
                      </p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/Boen.jpg' alt='Boen Liu' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Boen Liu"}</h5>
                    <p className='speaker-card-news__role'>{"Research Assistant"}</p>
                    <p className='speaker-card-news__org'>
                      {"HKUST(GZ) / MC"}
                      <sup>{"2"}</sup>
                      {" Lab"}
                    </p>
                    <p className='speaker-card-news__hint'>{"Click to view bio"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Research Assistant at MC² Lab, The Hong Kong University of Science and Technology (Guangzhou), and a student at Duke Kunshan University. His research focuses on social computing, graph learning and social network analysis."}</p>
                    </div>
                  </div>
                </article>
              </div>
              <h4>{"Three Hands-on Tracks"}</h4>
              <div className='track-grid'>
                <section className='track-card'>
                  <div className='track-card__head'>
                    <h5>{"The Art of Healing: Creative Arts as Pedagogical Practice"}</h5>
                  </div>
                  <div className='track-card__meta'>
                    <p className='track-card__lead'>
                      <strong>{"Leads:"}</strong>
                      <a href='#speaker-tengjia-zuo' className='agenda-speaker-link'>{"Tengjia Zuo"}</a>
                    </p>
                    <p className='track-card__time'>{"09:45 - 11:05 (80 minutes)"}</p>
                  </div>
                  <p className='track-card__body'>{"This session treats artistic practice not as an \"add-on\" to medical education, but as an essential way of knowing required for healing. Participants move through \"creative shifts\"—engaging in music, visual art, and narrative practices before translating those insights into educational design. Rather than producing polished curricula, the track focuses on creating pedagogical prototypes—concrete learning activities that embed creative disciplines into core medical training to cultivate capacities like deep listening, presence with suffering, and tolerance for ambiguity."}</p>
                </section>
                <section className='track-card'>
                  <div className='track-card__head'>
                    <h5>{"Health AI Systems Thinking for Community (HASTC)"}</h5>
                  </div>
                  <div className='track-card__meta'>
                    <p className='track-card__lead'>
                      <strong>{"Leads:"}</strong>
                      <a href='#speaker-mornin-feng' className='agenda-speaker-link'>{"Mornin Feng"}</a>
                    </p>
                    <p className='track-card__time'>{"11:15 - 12:45 (90 minutes)"}</p>
                  </div>
                  <p className='track-card__body'>{"This session encourages cross-disciplinary discussions and collaborative analyses of recent case studies regarding problems arising from clinical AI. Participants will critically review articles on algorithmic bias, model transparency, accountability, and unintended consequences in healthcare. Beyond identifying risks—such as data bias, hallucinations, and fairness limitations—the workshop asks teams to brainstorm and develop meaningful safeguards at the regulatory, institutional, and clinical levels to ensure technology does not exacerbate health inequities."}</p>
                </section>
                <section className='track-card'>
                  <div className='track-card__head'>
                    <h5>{"LLM-a-thon"}</h5>
                  </div>
                  <div className='track-card__meta'>
                    <p className='track-card__lead'>
                      <strong>{"Leads:"}</strong>
                      <a href='#speaker-calvin-kalun-or' className='agenda-speaker-link'>{"Calvin K.L. Or"}</a>
                    </p>
                    <p className='track-card__time'>{"15:55 - 17:15 (80 minutes)"}</p>
                  </div>
                  <p className='track-card__body'>{"This session invites participants to experiment with AI prompts and test the power of Large Language Models (LLMs) in clinical and mental health scenarios. Participants will compare responses across at least three different models to identify what is helpful, unhelpful, or risky in real-world use. A central goal is to experience LLM weaknesses firsthand, such as sycophancy bias, where systems may agree with users' leading questions rather than providing necessary medical guidance. The session culminates in an \"LLM as a Judge\" evaluation to critically analyze model safety, empathy, and advice quality."}</p>
                </section>
              </div>
              <h4>{"Venue"}</h4>
              <p>
                {"\n                        The workshop will be held at "}
                <strong>{"The Hong Kong University of Science and Technology (Guangzhou)"}</strong>
                {", 1 Duxue Road, Nansha District, Guangzhou, Guangdong, China. The event rooms are E4 1F 101 and E4 1F 102.\n                    "}
              </p>
              <p>{"\n                        By bringing healthcare, education, HCI, and AI governance into the same conversation, AI as Catalyst positions the campus as a space for practical experimentation, critical reflection, and cross-disciplinary collaboration around responsible AI.\n                    "}</p>
              <section className='sponsor-showcase sponsor-showcase--after-venue' aria-label='Workshop sponsors'>
                <div className='sponsor-showcase__intro'>
                  <span className='sponsor-showcase__eyebrow'>{"With Thanks to Our Sponsors"}</span>
                  <p className='sponsor-showcase__deck'>{"We gratefully acknowledge two technology companies whose support helps make this workshop possible."}</p>
                </div>
                <div className='sponsor-grid'>
                  <article className='sponsor-card'>
                    <div className='sponsor-card__header'>
                      <div className='sponsor-card__logo-frame'>
                        <img src='/images/workshop-sponsors/qingsong-health-logo.jpg' alt='QingSong Health Corporation logo' className='sponsor-card__logo sponsor-card__logo--qingsong' />
                      </div>
                      <div className='sponsor-card__heading'>
                        <h5>{"QingSong Health Corporation"}</h5>
                        <p className='sponsor-card__lead'>{"Founded in 2014, QingSong Health Corporation (HKEX: 02661.HK) is a leading one-stop health technology platform in China, specializing in integrated healthcare services and health insurance solutions. Leveraging its strengths in health education scenarios, industrial ecosystem synergy, and a continuously evolving technical framework, the Group builds a digital health service system driven by Artificial Intelligence to address users' needs across their entire life cycle."}</p>
                      </div>
                    </div>
                    <p>{"In terms of AI capability building, QingSong Health Corporation has established a full-stack technical ecosystem centered on its proprietary technology stack \"Alcare\" and its foundational medical health large language model \"Dr.GPT.\" This system covers critical scenarios such as health literacy, medical research, clinical decision-making, insurance services, and corporate health management, driving the evolution of health services from \"information connectivity\" to \"intelligent decision-making.\""}</p>
                    <p>{"Recently, the Group launched the evidence-based medicine AI agent \"ZhengYuanFang\" and unveiled the industry's first AI skills store, the ZhengYuanFang MedClaw Skills Store. By integrating massive volumes of authoritative medical literature and clinical guidelines, the platform upgrades evidence-based medicine from a traditional \"expert-experience-driven\" model to an \"intelligent-invocation-driven\" model. This provides doctors with more efficient and traceable clinical decision support, steering the healthcare service system toward greater standardization and intelligence."}</p>
                  </article>
                  <article className='sponsor-card' id='sponsor-wsc'>
                    <div className='sponsor-card__header'>
                      <div className='sponsor-card__logo-frame'>
                        <img src='/images/workshop-sponsors/wsc-holding-logo.jpg' alt='WSC HOLDING LIMITED logo' className='sponsor-card__logo sponsor-card__logo--wsc' />
                      </div>
                      <div className='sponsor-card__heading'>
                        <h5>{"WSC HOLDING LIMITED"}</h5>
                        <p className='sponsor-card__lead'>{"WSC HOLDING LIMITED is a dynamic biotechnology startup driven by a grand vision to revolutionize global healthcare. We are committed to developing innovative solutions across two strategic pillars: pioneering dental products and a holistic healthy aging platform. By forging strategic partnerships with world-renowned universities, we leverage cutting-edge biotechnology to create precise, high-efficiency medical solutions aimed at elevating human health standards and enhancing quality of life."}</p>
                      </div>
                    </div>
                    <p>{"Our dental project is currently in the early feasibility research stage, with the goal of developing a breakthrough product that redefines industry standards. Concurrently, our integrated healthy aging platform empowers individuals to maintain cardiovascular and bone health while optimizing daily habits through continuous monitoring and multi-modal interventions. Our mission is to enable people to remain vibrant throughout their lives, meeting the sophisticated demands for refined health management in the new consumer era."}</p>
                  </article>
                </div>
              </section>
              <h4>{"Registration"}</h4>
              <div className='registration-panel'>
                <div className='registration-panel__content'>
                  <p className='registration-panel__lead'>{"Please complete the registration form if you plan to attend the workshop."}</p>
                  <a className='registration-panel__link' href='https://forms.gle/h9GphSJ8dzZkESZ17' target='_blank' rel='noopener noreferrer'>{"Open registration form"}</a>
                  <p className='registration-panel__subtext'>{"You can also scan the QR code to register on mobile."}</p>
                </div>
                <a className='registration-panel__qr' href='https://forms.gle/h9GphSJ8dzZkESZ17' target='_blank' rel='noopener noreferrer' aria-label='Open the registration form'>
                  <img src='/images/workshop/ai_as_catalyst_registration_qr.png' alt='QR code for the AI as Catalyst workshop registration form.' />
                </a>
              </div>
              <a className='back-news-link' href='/news' target='_top' rel='noopener noreferrer'>{"← Back to News"}</a>
            </article>
          </div>
        </div>
      </div>
    </div>
  );
}

function ChineseWorkshopArticle(): JSX.Element {
  return (
    <div className='news-section'>
      <div className='container'>
        <div className='row justify-content-center'>
          <div className='col-12 text-center' style={{ paddingTop: '100px', color: '#fff' }}>
            <h3>{"新闻"}</h3>
          </div>
          <div className='col-lg-10 mt-4 mb-5'>
            <article className='news-article-card'>
              <h2 className='article-title'>{"AI as Catalyst 工作坊将在香港科技大学（广州）聚焦 AI 与医疗、教育的交汇"}</h2>
              <p className='article-meta'>
                <strong>{"发布日期："}</strong>
                {"2026年4月10日"}
              </p>
              <section className='affiliation-showcase' aria-label='工作坊组织机构与参与机构'>
                <div className='affiliation-group'>
                  <p className='affiliation-group__title'>{"组织机构"}</p>
                  <div className='organizer-lockup'>
                    <div className='organizer-lockup__row organizer-lockup__row--single'>
                      <div className='organizer-pair organizer-pair--schools'>
                        <div className='organizer-brand organizer-brand--mit'>
                          <div className='organizer-brand__visual'>
                            <img src='https://1000logos.net/wp-content/uploads/2026/01/MIT-logo.png' alt='麻省理工学院标识' className='organizer-brand__logo organizer-brand__logo--mit' />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"麻省理工学院"}</span>
                            <span className='organizer-brand__en'>{"MIT"}</span>
                          </div>
                        </div>
                        <span className='organizer-pair__mark' aria-hidden='true'>{"×"}</span>
                        <div className='organizer-brand organizer-brand--hkustgz'>
                          <div className='organizer-brand__visual'>
                            <img src='/images/UST-GZ-EN.png' alt='香港科技大学（广州）标识' className='organizer-brand__logo organizer-brand__logo--hkustgz' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"香港科技大学（广州）"}</span>
                            <span className='organizer-brand__en'>{"HKUST(GZ)"}</span>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className='organizer-lockup__row organizer-lockup__row--single'>
                      <div className='organizer-pair organizer-pair--institutions'>
                        <div className='organizer-brand organizer-brand--critical-data'>
                          <div className='organizer-brand__visual'>
                            <img src='https://criticaldatathon.github.io/static/img/logo.svg' alt='MIT Critical Data 标识' className='organizer-brand__logo organizer-brand__logo--critical-data' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"MIT 关键数据项目"}</span>
                            <span className='organizer-brand__en'>{"MIT Critical Data"}</span>
                          </div>
                        </div>
                        <span className='organizer-pair__mark' aria-hidden='true'>{"×"}</span>
                        <div className='organizer-brand organizer-brand--cma'>
                          <div className='organizer-brand__visual'>
                            <img src='/images/workshop-logos/cma-official.png' alt='计算媒体与艺术学域标识' className='organizer-brand__logo organizer-brand__logo--cma' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>{"计算媒体与艺术学域"}</span>
                            <span className='organizer-brand__en'>{"Computational Media and Arts"}</span>
                          </div>
                        </div>
                        <span className='organizer-pair__mark' aria-hidden='true'>{"×"}</span>
                        <div className='organizer-brand organizer-brand--mc2'>
                          <div className='organizer-brand__visual'>
                            <img src='/images/MC2.png' alt='元宇宙与计算创意中心（MC²）标识' className='organizer-brand__logo organizer-brand__logo--mc2' onError={hideOrganizerBrandImage} />
                          </div>
                          <div className='organizer-brand__meta'>
                            <span className='organizer-brand__zh'>
                              {"元宇宙与计算创意中心（MC"}
                              <sup>{"2"}</sup>
                              {"）"}
                            </span>
                            <span className='organizer-brand__en'>
                              {"MC"}
                              <sup>{"2"}</sup>
                              {" Lab"}
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
                <div className='affiliation-group affiliation-group--is-hidden'>
                  <p className='affiliation-group__title'>{"参与机构"}</p>
                  <div className='affiliation-grid affiliation-grid--participants'>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--dark'>
                        <img src='/images/UST-EN.png' alt='香港科技大学校徽' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"香港科技大学"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/harvard.png' alt='哈佛大学校徽' className='affiliation-tile__logo affiliation-tile__logo--participant affiliation-tile__logo--participant-crest affiliation-tile__logo--participant-harvard' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"哈佛大学"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/nus.png' alt='新加坡国立大学校徽' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"新加坡国立大学"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/hku.png' alt='香港大学校徽' className='affiliation-tile__logo affiliation-tile__logo--participant affiliation-tile__logo--participant-crest' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"香港大学"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/cuhk.png' alt='香港中文大学校徽' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"香港中文大学"}</p>
                    </article>
                    <article className='affiliation-tile'>
                      <div className='affiliation-tile__media affiliation-tile__media--participant'>
                        <img src='/images/workshop-logos/emory.png' alt='埃默里大学校徽' className='affiliation-tile__logo affiliation-tile__logo--participant' onError={hideAffiliationTileImage} />
                      </div>
                      <p className='affiliation-tile__name'>{"埃默里大学"}</p>
                    </article>
                  </div>
                </div>
              </section>
              <img className='article-cover' src='/images/20260410_ai_as_catalyst_workshop_cover.jpg' alt='AI as Catalyst 工作坊香港科技大学（广州）校园配图。' onError={useFallbackPhoto} />
              <div className='article-info-grid'>
                <div className='article-info-box'>
                  <span className='info-label'>{"活动日期"}</span>
                  <span className='info-value'>{"2026年5月12日"}</span>
                </div>
                <div className='article-info-box'>
                  <span className='info-label'>{"地点"}</span>
                  <span className='info-value'>{"中国广州"}</span>
                </div>
                <div className='article-info-box'>
                  <span className='info-label'>{"形式"}</span>
                  <span className='info-value'>{"全天工作坊与圆桌讨论"}</span>
                </div>
              </div>
              <p>
                <strong>{"AI as Catalyst"}</strong>
                {" 将于 2026 年 5 月 12 日在香港科技大学（广州）举行。这是一场面向学生、教育者、临床工作者与研究者的全天工作坊，聚焦人工智能在医疗与教育中的设计、评估、应用与治理。\n                    "}
              </p>
              <p>{"\n                        整场活动将动手实验、跨学科讨论与协作式设计结合在一起。它并不把 AI 视作单纯提升效率的工具，而是强调批判性参与：参与者既要测试能力边界，也要识别失效情境，并进一步讨论在真实部署前所需的社会与制度保障。\n                    "}</p>
              <h4>{"议程速览"}</h4>
              <div className='agenda-snapshot'>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"09:00-09:30"}</span>
                  <div>
                    <strong>{"签到与欢迎"}</strong>
                    <p>{"签到入座。"}</p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"09:30-09:45"}</span>
                  <div>
                    <strong>{"开幕式"}</strong>
                    <p>
                      {"工作坊概览 / 欢迎致辞。主持："}
                      <a href='#speaker-pan-hui' className='agenda-speaker-link'>{"Pan Hui"}</a>
                      {" 与 "}
                      <a href='#speaker-leo-anthony-celi' className='agenda-speaker-link'>{"Leo Anthony Celi"}</a>
                      {"。"}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"09:45-11:05"}</span>
                  <div>
                    <strong>{"工作坊一"}</strong>
                    <p>
                      {"治愈的艺术：作为教学实践的创意艺术。主讲："}
                      <a href='#speaker-tengjia-zuo' className='agenda-speaker-link'>{"Tengjia Zuo"}</a>
                      {"。"}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"11:05-11:15"}</span>
                  <div>
                    <strong>{"中场休息"}</strong>
                    <p>{"交流。"}</p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"11:15-12:45"}</span>
                  <div>
                    <strong>{"工作坊二"}</strong>
                    <p>
                      {"面向社群的健康 AI 系统思维（HASTC）。主讲："}
                      <a href='#speaker-mornin-feng' className='agenda-speaker-link'>{"Mornin Feng"}</a>
                      {"。"}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"12:45-14:35"}</span>
                  <div>
                    <strong>
                      {"午餐与"}
                      <a href='https://surreality.world/' target='_blank' rel='noopener noreferrer'>
                        <em>{"Surreality元宇宙艺术展"}</em>
                      </a>
                      {"参观"}
                    </strong>
                    <p>{"午餐与艺术展参观。"}</p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"14:35-15:35"}</span>
                  <div>
                    <strong>{"座谈会"}</strong>
                    <p>
                      {"座谈题目：从个人能动性到系统性安全：医疗领域的真实 AI 实践。主持："}
                      <a href='#speaker-qiushi-zhou' className='agenda-speaker-link'>{"Qiushi Zhou"}</a>
                      {"。"}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"15:35-15:55"}</span>
                  <div>
                    <strong>
                      <a href='#sponsor-wsc' className='agenda-speaker-link'>{"WSC 冠名茶歇"}</a>
                    </strong>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"15:55-17:15"}</span>
                  <div>
                    <strong>{"工作坊三"}</strong>
                    <p>
                      {"LLM-a-thon：从使用场景到事实校验。主讲："}
                      <a href='#speaker-calvin-kalun-or' className='agenda-speaker-link'>{"Calvin K.L. Or"}</a>
                      {"。"}
                    </p>
                  </div>
                </div>
                <div className='agenda-snapshot__item'>
                  <span className='agenda-snapshot__time'>{"17:15-17:30"}</span>
                  <div>
                    <strong>{"闭幕式"}</strong>
                    <p>
                      {"关键收获 / 后续安排 / 总结。主持："}
                      <a href='#speaker-pan-hui' className='agenda-speaker-link'>{"Pan Hui"}</a>
                      {" 与 "}
                      <a href='#speaker-leo-anthony-celi' className='agenda-speaker-link'>{"Leo Anthony Celi"}</a>
                      {"。"}
                    </p>
                  </div>
                </div>
              </div>
              <h4>{"主讲人与组织者"}</h4>
              <p>{"\n                        本次工作坊邀请了来自香港科技大学（广州）、香港科技大学、麻省理工学院、哈佛大学、新加坡国立大学、香港大学、香港中文大学、埃默里大学及相关合作机构的主讲人与组织者。点击任意卡片即可展开查看人物简介。\n                    "}</p>
              <div className='speaker-grid-news'>
                <article className='speaker-card-news' id='speaker-pan-hui' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/pan_hui.jpg?v=20260412-2' alt='Pan Hui' className='speaker-card-news__image speaker-card-news__image--pan-hui' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Pan Hui"}</h5>
                    <p className='speaker-card-news__role'>{"讲席教授"}</p>
                    <p className='speaker-card-news__org'>
                      {"香港科技大学（广州） / 元宇宙与计算创意中心（MC"}
                      <sup>{"2"}</sup>
                      {"）"}
                    </p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"英国皇家工程院国际院士、IEEE Fellow、ACM Distinguished Scientist、欧洲科学院院士，现任香港科技大学（广州）计算媒体与艺术讲席教授。他的研究聚焦于普适计算、移动计算、增强/虚拟现实、数据科学、社交与移动网络，以及元宇宙相关系统。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-leo-anthony-celi' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/leo_celi.jpg' alt='Leo Anthony Celi' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Leo Anthony Celi"}</h5>
                    <p className='speaker-card-news__role'>{"高级研究科学家 / 副教授"}</p>
                    <p className='speaker-card-news__org'>{"麻省理工学院 / 哈佛大学"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任麻省理工学院高级研究科学家、MIT 计算生理学实验室临床研究主任、MIT Sana 联合主任、Beth Israel Deaconess Medical Center 肺与危重症科医生，并兼任哈佛医学院医学副教授。他的工作聚焦于利用数据科学与人工智能改进重症医疗，尤其关注开放临床数据以及减少医疗系统中的偏差。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/Chenlei.png' alt='Lei Chen' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Lei Chen"}</h5>
                    <p className='speaker-card-news__role'>{"信息枢纽院长 / 讲席教授"}</p>
                    <p className='speaker-card-news__org'>{"香港科技大学（广州） / 信息枢纽"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"陈雷现任香港科技大学（广州）信息枢纽院长、香港科技大学大数据研究所所长，并在数据科学与分析学域及人工智能学域担任讲席教授。他是 ACM Fellow、IEEE Fellow。其研究兴趣包括数据驱动的机器学习、基于众包的数据处理、不确定与概率数据库、Web 信息管理、多媒体系统、知识图谱、区块链、数据隐私以及时空数据管理。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-mornin-feng' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/mornin_feng.jpg' alt='Mornin Feng' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Mornin Feng"}</h5>
                    <p className='speaker-card-news__role'>{"副教授"}</p>
                    <p className='speaker-card-news__org'>{"新加坡国立大学"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任新加坡国立大学 Saw Swee Hock 公共卫生学院副教授、AI for Public Health（AI4PH）项目主任、生物统计、建模、AI 与数据分析（B.MAD）领域负责人，并兼任杨潞龄医学院、生物医学工程系与数据科学研究所副教授，以及 OHDSI 新加坡分会主席。他的研究聚焦于健康医疗 AI、临床数据机器学习与数据驱动决策，应用涵盖医学影像、治疗优化与临床自然语言处理。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-calvin-kalun-or' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/calvin_or.jpg' alt='Calvin Kalun Or' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Calvin Kalun Or"}</h5>
                    <p className='speaker-card-news__role'>{"副教授"}</p>
                    <p className='speaker-card-news__org'>{"香港大学"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任香港大学数据与系统工程学系副教授兼副系主任。他的研究聚焦于人因工程、人机交互与医疗信息学，特别关注如何通过健康信息技术的设计、实施与评估提升患者安全、系统表现与医疗服务质量。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/Lowell_Ling.jpg' alt='Lowell Ling' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Lowell Ling"}</h5>
                    <p className='speaker-card-news__role'>{"副教授"}</p>
                    <p className='speaker-card-news__org'>{"香港中文大学"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任香港中文大学麻醉与深切治疗学系副教授，并于威尔斯亲王医院深切治疗部担任助理教授。他的研究聚焦于脓毒症、器官功能障碍与重症医学，尤其关注香港地区脓毒症的流行病学及其基因组机制。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/xiangliangzhang.jpg' alt='Nancy Zhang' className='speaker-card-news__image speaker-card-news__image--xianglilan' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Nancy Zhang"}</h5>
                    <p className='speaker-card-news__role'>{"助理教授"}</p>
                    <p className='speaker-card-news__org'>{"香港中文大学"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任香港中文大学医学院助理教授。她的研究聚焦于将人工智能与机器学习应用于临床诊断、疾病监测与治疗决策支持，在精准医疗和重症医学等方向开展研究。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-qiushi-zhou' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/qiushi_zhou_20260423.png' alt='Qiushi Zhou' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Qiushi Zhou"}</h5>
                    <p className='speaker-card-news__role'>{"助理教授"}</p>
                    <p className='speaker-card-news__org'>{"香港科技大学（广州）"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任香港科技大学（广州）助理教授，隶属计算媒体与艺术学域及物联网学域。他的研究聚焦于人机交互与扩展现实，尤其关注 XR、AI 与 IoT 融合下的新型交互技术。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/xin_tong.jpg' alt='Xin Tong' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Xin Tong"}</h5>
                    <p className='speaker-card-news__role'>{"助理教授"}</p>
                    <p className='speaker-card-news__org'>{"香港科技大学（广州）"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任香港科技大学（广州）计算媒体与艺术学域助理教授。她主要从事人机交互与人机协作研究，关注医疗、福祉与无障碍等议题，并探索 VR/AR 与生成式 AI 等技术在这些场景中的应用。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' id='speaker-tengjia-zuo' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/tengjia_zuo.jpg' alt='Tengjia Zuo' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Tengjia Zuo"}</h5>
                    <p className='speaker-card-news__role'>{"助理教授"}</p>
                    <p className='speaker-card-news__org'>{"香港科技大学（广州）"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任香港科技大学（广州）计算媒体与艺术学域助理教授。她的研究聚焦于混合现实、严肃游戏与人机交互，特别关注游戏化学习与玩家体验设计。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/Hyunjung.jpg' alt='Hyunjung Gloria Kwak' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Hyunjung Gloria Kwak"}</h5>
                    <p className='speaker-card-news__role'>{"助理教授"}</p>
                    <p className='speaker-card-news__org'>{"埃默里大学"}</p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现任埃默里大学护理学院助理教授。她的研究聚焦于利用人工智能与机器学习提升临床决策质量，尤其关注公平性推进与医疗健康差异的缩减。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/james_formal.jpg' alt='James Yiming Zhu' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"James Yiming Zhu"}</h5>
                    <p className='speaker-card-news__role'>{"博士生"}</p>
                    <p className='speaker-card-news__org'>
                      {"香港科技大学（广州） / 元宇宙与计算创意中心（MC"}
                      <sup>{"2"}</sup>
                      {"）"}
                    </p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现为香港科技大学（广州）博士生，并隶属于元宇宙与计算创意中心（MC²）。他的研究兴趣主要包括社会计算、数据科学与自然语言处理。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/yulin_yao.jpg' alt='Yulin Yao' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Yulin Yao"}</h5>
                    <p className='speaker-card-news__role'>{"博士生 / 视觉艺术家"}</p>
                    <p className='speaker-card-news__org'>
                      {"香港科技大学（广州） / 元宇宙与计算创意中心（MC"}
                      <sup>{"2"}</sup>
                      {"）"}
                    </p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Yulin Yao 是香港科技大学（广州）计算媒体与艺术方向博士生，同时也是一名数字视觉艺术家。她的工作位于沉浸式叙事、VR 与 HCI 的交叉处。基于其在欧洲与中国长期积累的空间设计和展览实践，她创作沉浸式数字环境，探索技术如何影响记忆、情感与文化。她的实践连接计算媒体与视觉艺术，并将空间本身视为心理探询与沉浸式世界构建的媒介。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/bianca_yang.jpg' alt='Bianca Ruoshan Yang' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Bianca Ruoshan Yang"}</h5>
                    <p className='speaker-card-news__role'>{"博士生 / 音乐人"}</p>
                    <p className='speaker-card-news__org'>
                      {"香港科技大学（广州） / 元宇宙与计算创意中心（MC"}
                      <sup>{"2"}</sup>
                      {"）"}
                    </p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"Bianca Ruoshan Yang 是香港科技大学（广州）博士生，隶属于元宇宙与计算创意中心（MC²）。她的研究位于多模态 AI、以学习为中心的 XR 设计，以及非物质文化遗产的交叉地带。"}</p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/Team_Profile_Pic/primo_formal.jpeg' alt='Primo Dongyijie Pan' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Primo Dongyijie Pan"}</h5>
                    <p className='speaker-card-news__role'>{"MPhil Student"}</p>
                    <p className='speaker-card-news__org'>
                      {"香港科技大学（广州） / 元宇宙与计算创意中心（MC"}
                      <sup>{"2"}</sup>
                      {"）"}
                    </p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>
                        {"Primo 的研究关注如何将现有心理治疗技术，尤其是认知行为疗法，转译为人机交互的语言。他参与开发的 Lingolift、Virtual AI-SP 等项目已在 CHI 等会议发表，并获得中国计算机学会和腾讯 Light 基金会的认可与奖励。面向博士阶段的研究，他希望推动 AI、XR 等技术与前沿生活干预方法结合，为慢病管理带来更多可能。个人主页："}
                        <a href='https://primopan.github.io/about/' target='_blank' rel='noopener noreferrer'>{"https://primopan.github.io/about/"}</a>
                      </p>
                    </div>
                  </div>
                </article>
                <article className='speaker-card-news' tabIndex={0} role='button' aria-expanded='false'>
                  <img src='/images/workshop-speakers/Boen.jpg' alt='Boen Liu' className='speaker-card-news__image' onError={useFallbackPhoto} />
                  <div className='speaker-card-news__body'>
                    <h5>{"Boen Liu"}</h5>
                    <p className='speaker-card-news__role'>{"Research Assistant"}</p>
                    <p className='speaker-card-news__org'>
                      {"香港科技大学（广州） / 元宇宙与计算创意中心（MC"}
                      <sup>{"2"}</sup>
                      {"）"}
                    </p>
                    <p className='speaker-card-news__hint'>{"点击查看简介"}</p>
                    <div className='speaker-card-news__details'>
                      <p className='speaker-card-news__bio'>{"现于香港科技大学（广州）元宇宙与计算创意中心（MC²）担任研究助理，并就读于昆山杜克大学。他的研究聚焦于社会计算、图学习与社交网络分析。"}</p>
                    </div>
                  </div>
                </article>
              </div>
              <h4>{"三条实践工作坊主线"}</h4>
              <div className='track-grid'>
                <section className='track-card'>
                  <div className='track-card__head'>
                    <h5>{"治愈的艺术：创意艺术作为教学实践"}</h5>
                  </div>
                  <div className='track-card__meta'>
                    <p className='track-card__lead'>
                      <strong>{"主讲："}</strong>
                      <a href='#speaker-tengjia-zuo' className='agenda-speaker-link'>{"Tengjia Zuo"}</a>
                    </p>
                    <p className='track-card__time'>{"09:45 - 11:05（80分钟）"}</p>
                  </div>
                  <p className='track-card__body'>{"本场工作坊并不将艺术实践视为医学教育中的“附加项”，而是将其理解为疗愈所必需的一种根本性认知方式。参与者将经历一系列“创造性转向”——先投入音乐、视觉艺术与叙事实践，再把这些体验与洞见转化为教育设计。与其着手打磨一套完整成熟的课程方案，这一环节更强调产出教学原型：即把创意学科嵌入核心医学训练的具体学习活动，以培养深度倾听、与痛苦同在，以及面对不确定性的能力。"}</p>
                </section>
                <section className='track-card'>
                  <div className='track-card__head'>
                    <h5>{"面向社群的健康 AI 系统思维（HASTC）"}</h5>
                  </div>
                  <div className='track-card__meta'>
                    <p className='track-card__lead'>
                      <strong>{"主讲："}</strong>
                      <a href='#speaker-mornin-feng' className='agenda-speaker-link'>{"Mornin Feng"}</a>
                    </p>
                    <p className='track-card__time'>{"11:15 - 12:45（90分钟）"}</p>
                  </div>
                  <p className='track-card__body'>{"本场工作坊鼓励围绕近期临床 AI 案例中暴露的问题，展开跨学科讨论与协作分析。参与者将批判性审阅有关算法偏见、模型透明性、问责机制以及医疗场景中非预期后果的相关文章。除了识别风险——例如数据偏差、模型幻觉与公平性局限——工作坊还将邀请各小组从监管、机构与临床多个层面共同构思并提出切实可行的防护机制，确保技术不会进一步加剧健康不平等。"}</p>
                </section>
                <section className='track-card'>
                  <div className='track-card__head'>
                    <h5>{"LLM-a-thon"}</h5>
                  </div>
                  <div className='track-card__meta'>
                    <p className='track-card__lead'>
                      <strong>{"主讲："}</strong>
                      <a href='#speaker-calvin-kalun-or' className='agenda-speaker-link'>{"Calvin K.L. Or"}</a>
                    </p>
                    <p className='track-card__time'>{"15:55 - 17:15（80分钟）"}</p>
                  </div>
                  <p className='track-card__body'>{"本场工作坊邀请参与者围绕 AI 提示词展开实验，并在临床与心理健康情境中测试大语言模型（LLMs）的能力。参与者将比较至少三种不同模型的回答，以识别哪些输出在真实应用中有帮助、无帮助或存在风险。一个核心目标，是让参与者亲身体验 LLM 的局限性，例如“迎合性偏差”（sycophancy bias）：系统可能顺着用户带有倾向性的提问附和，而不是提供必要的医疗指引。最后，本场环节将以“LLM as a Judge”评估收束，从模型安全性、共情能力与建议质量等维度展开批判性分析。"}</p>
                </section>
              </div>
              <h4>{"举办地点"}</h4>
              <p>
                {"\n                        活动将在"}
                <strong>{"香港科技大学（广州）"}</strong>
                {"举行，地址为中国广东省广州市南沙区笃学路1号，具体活动教室为 E4 1F 101 和 E4 1F 102。\n                    "}
              </p>
              <p>{"\n                        通过把医疗、教育、人机交互与 AI 治理放进同一个讨论现场，AI as Catalyst 也把校园变成了一个用于实践检验、批判反思与跨学科协作的空间，让“负责任 AI”不只停留在口号层面，而进入具体方法与真实场景。\n                    "}</p>
              <section className='sponsor-showcase sponsor-showcase--after-venue' aria-label='工作坊赞助支持'>
                <div className='sponsor-showcase__intro'>
                  <span className='sponsor-showcase__eyebrow'>{"鸣谢赞助支持"}</span>
                  <p className='sponsor-showcase__deck'>{"感谢两家科技企业对本次工作坊的支持，使这一跨学科交流与实践活动得以顺利开展。"}</p>
                </div>
                <div className='sponsor-grid'>
                  <article className='sponsor-card'>
                    <div className='sponsor-card__header'>
                      <div className='sponsor-card__logo-frame'>
                        <img src='/images/workshop-sponsors/qingsong-health-logo.jpg' alt='轻松健康集团标识' className='sponsor-card__logo sponsor-card__logo--qingsong' />
                      </div>
                      <div className='sponsor-card__heading'>
                        <h5>{"轻松健康集团"}</h5>
                        <p className='sponsor-card__lead'>{"轻松健康集团（港股代码：02661.HK）成立于 2014 年，是国内领先的一站式健康科技平台，专注于提供综合健康服务与健康保险解决方案。依托健康教育场景优势、产业生态协同能力与持续迭代的技术体系，集团围绕用户全生命周期健康需求，持续构建以人工智能为核心驱动力的数字健康服务体系。"}</p>
                      </div>
                    </div>
                    <p>{"在 AI 能力建设方面，轻松健康集团形成了以自研技术栈“Alcare”和医疗健康基座大模型“Dr.GPT”为核心的全链路技术体系，覆盖健康科普、医学研究、临床决策、保险服务及企业健康管理等多个关键场景，推动健康服务从“信息连接”向“智能决策”升级。"}</p>
                    <p>{"近期，集团推出循证医学智能体“证元芳”，并发布医疗行业首个 AI 技能商店——证元芳 MedClaw Skills Store。通过整合海量权威医学文献与临床指南，将循证医学能力从传统“专家经验驱动”，升级为“智能调用驱动”，为医生提供更加高效、可溯源的临床决策支持，推动医疗服务体系向标准化、智能化迈进。"}</p>
                  </article>
                  <article className='sponsor-card' id='sponsor-wsc'>
                    <div className='sponsor-card__header'>
                      <div className='sponsor-card__logo-frame'>
                        <img src='/images/workshop-sponsors/wsc-holding-logo.jpg' alt='WSC控股有限公司标识' className='sponsor-card__logo sponsor-card__logo--wsc' />
                      </div>
                      <div className='sponsor-card__heading'>
                        <h5>{"WSC控股有限公司（WSC HOLDING LIMITED）"}</h5>
                        <p className='sponsor-card__lead'>{"WSC控股有限公司（WSC HOLDING LIMITED）是一家充满活力的生物科技初创公司，怀揣着为全球医疗健康事业带来革新性改变的宏伟愿景，致力于在两大关键领域开发创新解决方案：创新牙科产品和整体健康老龄化平台。我们正通过与世界知名大学的战略合作推进这些重点项目，运用前沿的生物技术，开发精准、高效的医疗解决方案，以提升人类健康水准，改善人们生活品质。"}</p>
                      </div>
                    </div>
                    <p>{"我们的牙科项目目前正处于早期可行性研究阶段，目标是开发一款突破性的新产品。与此同时，我们的综合健康老龄化平台通过持续监测和多模式干预，帮助人们维护心血管与骨骼健康、优化日常习惯，让人们能够终身保持活力，满足新消费时代下精细化健康管理需求。"}</p>
                  </article>
                </div>
              </section>
              <h4>{"报名方式"}</h4>
              <div className='registration-panel'>
                <div className='registration-panel__content'>
                  <p className='registration-panel__lead'>{"如计划参加本次工作坊，请填写报名问卷完成登记。"}</p>
                  <a className='registration-panel__link' href='https://forms.gle/h9GphSJ8dzZkESZ17' target='_blank' rel='noopener noreferrer'>{"打开报名问卷"}</a>
                  <p className='registration-panel__subtext'>{"也可以直接扫描二维码，在手机上完成报名。"}</p>
                </div>
                <a className='registration-panel__qr' href='https://forms.gle/h9GphSJ8dzZkESZ17' target='_blank' rel='noopener noreferrer' aria-label='打开报名问卷'>
                  <img src='/images/workshop/ai_as_catalyst_registration_qr.png' alt='AI as Catalyst 工作坊报名问卷二维码。' />
                </a>
              </div>
              <a className='back-news-link' href='/zh/news' target='_top' rel='noopener noreferrer'>{"← 返回新闻列表"}</a>
            </article>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function NewsAIAsCatalystWorkshopPage({ locale = 'en' }: NewsAIAsCatalystWorkshopPageProps): JSX.Element {
  const pageRef = useRef<HTMLElement>(null);
  useWorkshopSpeakerCards(pageRef);
  const isZh = locale === 'zh-CN';
  const alternateHref = isZh ? '/news/ai-as-catalyst-workshop' : '/zh/news/ai-as-catalyst-workshop';

  return (
    <SitePageShell className='workshop-page-shell' ariaLabel={isZh ? 'AI as Catalyst 工作坊' : 'AI as Catalyst Workshop'} locale={locale} activeRoute='news' alternateHref={alternateHref} ref={pageRef}>
      {isZh ? <ChineseWorkshopArticle /> : <EnglishWorkshopArticle />}
    </SitePageShell>
  );
}
