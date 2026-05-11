import React, { useEffect, useMemo, useRef, useState } from 'react';
import { recentPublicationData } from '../data/recentPublications';
import '../styles/publication-page.css';

interface PublicationPageProps {
  locale?: 'en' | 'zh-CN';
}

const FALLBACK_PHOTO = 'http://www.ivang-design.com/svg-load/portfolio/photo-p.jpg';
const YEARS = ['2026', '2025', '2024', '2023', '2022', '2021'];

const navItems = {
  en: [
    { href: '/', label: 'Home' },
    { href: '/people', label: 'People' },
    { href: '/publication', label: 'Publication', active: true },
    { href: '/project', label: 'Project' },
    { href: '/news', label: 'News' },
    { href: '/leader', label: 'Director' }
  ],
  zh: [
    { href: '/zh', label: '首页' },
    { href: '/zh/people', label: '成员' },
    { href: '/zh/publication', label: '论文', active: true },
    { href: '/zh/project', label: '项目' },
    { href: '/zh/news', label: '新闻' },
    { href: '/zh/leader', label: '负责人' }
  ]
};

function useFallbackPhoto(event: React.SyntheticEvent<HTMLImageElement>): void {
  const image = event.currentTarget;
  if (image.src === FALLBACK_PHOTO) return;
  image.src = FALLBACK_PHOTO;
}

function decoratePanHuiNames(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>('.author').forEach((element) => {
    if (element.innerHTML.includes('pan-hui-name')) return;
    const text = element.textContent || '';
    if (!text.includes('Pan Hui')) return;
    element.textContent = '';
    const parts = text.split('Pan Hui');
    parts.forEach((part, index) => {
      if (part) element.appendChild(document.createTextNode(part));
      if (index < parts.length - 1) {
        const span = document.createElement('span');
        span.className = 'pan-hui-name';
        span.textContent = 'Pan Hui';
        element.appendChild(span);
      }
    });
  });
}

function useBodyPublicationClass(): void {
  useEffect(() => {
    document.body.classList.add('publication-on');
    return () => document.body.classList.remove('publication-on');
  }, []);
}

function usePublicationScrollSpy(pageRef: React.RefObject<HTMLElement>, setActiveYear: React.Dispatch<React.SetStateAction<string>>): void {
  useEffect(() => {
    const page = pageRef.current;
    const scrollContainer = page?.querySelector<HTMLElement>('.publication-section');
    if (!page || !scrollContainer) return undefined;

    const syncYearNavByScroll = () => {
      const sections = YEARS
        .map((year) => page.querySelector<HTMLElement>(`#yr${year}`))
        .filter((section): section is HTMLElement => Boolean(section));
      if (!sections.length) return;

      const containerRect = scrollContainer.getBoundingClientRect();
      const triggerLine = scrollContainer.scrollTop + 140;
      let activeSection = sections[0];
      sections.forEach((section) => {
        const sectionTopInContainer = section.getBoundingClientRect().top - containerRect.top + scrollContainer.scrollTop;
        if (sectionTopInContainer <= triggerLine) activeSection = section;
      });
      setActiveYear(activeSection.id.replace('yr', ''));
    };

    scrollContainer.addEventListener('scroll', syncYearNavByScroll, { passive: true });
    window.addEventListener('resize', syncYearNavByScroll);
    syncYearNavByScroll();
    return () => {
      scrollContainer.removeEventListener('scroll', syncYearNavByScroll);
      window.removeEventListener('resize', syncYearNavByScroll);
    };
  }, [pageRef, setActiveYear]);
}

function PublicationNavBar({ locale }: { locale: 'en' | 'zh-CN' }): JSX.Element {
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const isZh = locale === 'zh-CN';
  const items = isZh ? navItems.zh : navItems.en;
  const alternateHref = isZh ? '/publication' : '/zh/publication';

  useEffect(() => {
    if (!isMenuOpen) return undefined;
    const closeMenu = (event: MouseEvent) => {
      const target = event.target;
      if (target instanceof Node && !document.querySelector('.major-nav')?.contains(target)) setIsMenuOpen(false);
    };
    document.addEventListener('click', closeMenu);
    return () => document.removeEventListener('click', closeMenu);
  }, [isMenuOpen]);

  return (
    <div className={isMenuOpen ? 'major-nav is-menu-open' : 'major-nav'}>
      <nav className='nav-logo'>
        <img src='/images/MC2.png' alt='MC2 Lab' />
        <div className='vertical'></div>
        <img src={isZh ? '/images/UST-GZ-ZH.png' : '/images/UST-GZ-EN.png'} alt={isZh ? '香港科技大学（广州）' : 'HKUST(GZ)'} />
        <div className='vertical'></div>
        <img src={isZh ? '/images/UST-ZH.png' : '/images/UST-EN.png'} alt={isZh ? '香港科技大学' : 'HKUST'} />
      </nav>
      <nav className='nav' style={{ margin: '0 .5em', padding: '0 .5em', boxSizing: 'content-box' }}>
        <ul className='nav__links' style={{ padding: '1em', paddingBottom: '0', gap: '30px', boxSizing: 'content-box' }}>
          {items.map((item) => (
            <li className={item.active ? 'nav__link active' : 'nav__link'} key={item.href}>
              <a href={item.href} style={{ fontFamily: 'Open Sans' }}>{item.label}</a>
            </li>
          ))}
        </ul>
      </nav>
      <div className='banner-controls'>
        <a className='banner-lang' href={alternateHref}>{isZh ? 'EN' : '中文'}</a>
        <button type='button' className='banner-menu-toggle' id='bannerMenuToggle' aria-expanded={isMenuOpen} aria-label={isZh ? '打开导航菜单' : 'Open navigation menu'} onClick={() => setIsMenuOpen((open) => !open)}>
          {isZh ? '菜单' : 'Menu'}
        </button>
      </div>
    </div>
  );
}

function YearNav({ activeYear, onSelectYear }: { activeYear: string; onSelectYear: (year: string) => void }): JSX.Element {
  return (
    <div className='publication-nav'>
      <div className='container'>
        <ul className='yr__navs'>
          {YEARS.map((year) => (
            <li className={activeYear === year ? `yr__nav yr${year} active` : `yr__nav yr${year}`} key={year}>
              <a href={`#yr${year}`} onClick={(event) => { event.preventDefault(); onSelectYear(year); }}>{year}</a>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}

function RecentPublicationItem({ title, authors, venue }: { title: string; authors: string; venue: string }): JSX.Element {
  return (
    <article className='recent-publication-item'>
      <h4>{title}</h4>
      <p className='conference'>{venue}</p>
      <p className='author'>{authors}</p>
    </article>
  );
}

function RecentPublicationSections(): JSX.Element {
  const years = useMemo(() => Object.keys(recentPublicationData).sort((a, b) => Number(b) - Number(a)), []);
  return (
    <>
      {years.map((year) => (
        <section id={`yr${year}`} className='recent-year-section' key={year}>
          <br />
          <br />
          <h3>{`YEAR ${year}`}</h3>
          <div className='recent-publication-list'>
            {recentPublicationData[year].map((entry, index) => <RecentPublicationItem {...entry} key={`${year}-${index}-${entry.title}`} />)}
          </div>
        </section>
      ))}
    </>
  );
}

function StaticPublicationSections(): JSX.Element {
  return (
    <>
      <section id={"yr2024"}>
        <br />
        <br />
        <h3>{"YEAR 2024"}</h3>
        <div className={"row"}>
          <div className={"col-md-4"}>
            <img src={"/images/Publication_png/SIGCHI_2024.png"} alt={"Cover image for Avatar Appearance and Behavior of Potential Harassers Affect Users? Perceptions and Response Strategies in Social Virtual Reality (VR): A Mixed-Methods Study."} onError={useFallbackPhoto} />
          </div>
          <div className={"col-md-8"}>
            <h4>{"Avatar Appearance and Behavior of Potential Harassers Affect Users? Perceptions and Response Strategies in Social Virtual Reality (VR): A Mixed-Methods Study."}</h4>
            <p className={"conference"}>{"27th ACM SIGCHI CSCW | November 2024"}</p>
            <p className={"author"}>{"Xuetong Wang, Ziyan Wang, Mingming Zhang, Kangyou Yu, Pan Hui, and Mingming Fan"}</p>
            <a href={""} className={"publication-btn"} target={"_blank"}>
              <i className={"fa fa-file-pdf"}></i>
              {" PDF"}
            </a>
            <a href={""} className={"publication-btn"} target={"_blank"}>
              <i className={"fa fa-video-camera"}></i>
              {" Video"}
            </a>
            <a href={"https://doi.org/None"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2407.10053.png"} alt={"Cover image for The Jade Gateway to Exergaming: How Socio-Cultural Factors Shape Exergaming Among East Asian Older Adults"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"The Jade Gateway to Exergaming: How Socio-Cultural Factors Shape Exergaming Among East Asian Older Adults"}</h4>
              <p className={"conference"}>{"ACM CHI Play | October 2024"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Juhyung Son, Simin Yang, Derrick M. Wang, Lydia Choong, Ahmad Alhilal, Pengyuan Zhou, Pan Hui, and Lennart E. Nacke"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2407.10053"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2407.05098.png"} alt={"Cover image for FedTSA: A Cluster-based Two-Stage Aggregation Method for Model-heterogeneous Federated Learning"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"FedTSA: A Cluster-based Two-Stage Aggregation Method for Model-heterogeneous Federated Learning"}</h4>
              <p className={"conference"}>{"ECCV | September 2024"}</p>
              <p className={"author"}>{"Boyu Fan, Chenrui Wu, Xiang Su, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2407.05098"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2407.06422.png"} alt={"Cover image for Exploring the Capability of ChatGPT to Reproduce Human Labels for Social Computing Tasks"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Exploring the Capability of ChatGPT to Reproduce Human Labels for Social Computing Tasks"}</h4>
              <p className={"conference"}>{"ASONAM | September 2024"}</p>
              <p className={"author"}>{"Yiming Zhu, Peixian Zhang, Ehsan-Ul Haq, Pan Hui, and Gareth Tyson"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2407.06422"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453637528.3671965.png"} alt={"Cover image for Large Language Model-driven Meta-structure Discovery in Heterogeneous Information Network"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Large Language Model-driven Meta-structure Discovery in Heterogeneous Information Network"}</h4>
              <p className={"conference"}>{"30th ACM SIGKDD | August 2024"}</p>
              <p className={"author"}>{"Lin Chen, Fengli Xu, Nian Li, Zhenyu Han, Meng Wang, Yong Li, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3637528.3671965"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.24963ijcai.2024797.png"} alt={"Cover image for VulnerabilityMap: An Open Framework for Mapping Vulnerability among Urban Disadvantaged Populations in the United States"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"VulnerabilityMap: An Open Framework for Mapping Vulnerability among Urban Disadvantaged Populations in the United States"}</h4>
              <p className={"conference"}>{"IJCAI | August 2024"}</p>
              <p className={"author"}>{"Lin Chen, Yong Li, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.24963/ijcai.2024/797"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2406.11282.png"} alt={"Cover image for From Pixels to Progress: Generating Road Network from Satellite Imagery for Socioeconomic Insights in Impoverished Areas"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"From Pixels to Progress: Generating Road Network from Satellite Imagery for Socioeconomic Insights in Impoverished Areas"}</h4>
              <p className={"conference"}>{"IJCAI | August 2024"}</p>
              <p className={"author"}>{"Yanxin Xi, Yu Liu, Zhicheng Liu, Sasu Tarkoma, Pan Hui,... and Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2406.11282"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2305.00510.png"} alt={"Cover image for Towards AI-Architecture Liberty: A Comprehensive Survey on Design and Generation of Virtual Architecture by Deep Learning"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards AI-Architecture Liberty: A Comprehensive Survey on Design and Generation of Virtual Architecture by Deep Learning"}</h4>
              <p className={"conference"}>{"ACM CSUR | July 2024"}</p>
              <p className={"author"}>{"Anqi Wang, Jiahua Dong, Lik-Hang Lee, Jiachuan Shen, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2305.00510"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1007978-3-031-60692-2_12.png"} alt={"Cover image for Making Learning Engaging and Productive: SimLab, A VR Lab to Bridge between Classroom Theory and Industrial Practice in Chemical Engineering Education"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Making Learning Engaging and Productive: SimLab, A VR Lab to Bridge between Classroom Theory and Industrial Practice in Chemical Engineering Education"}</h4>
              <p className={"conference"}>{"HCII | July 2024"}</p>
              <p className={"author"}>{"Juhyung Son, Ahmad Alhilal, Reza Hadi Mogavi, Tristan Braud, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1007/978-3-031-60692-2_12"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1007978-3-031-60692-2_3.png"} alt={"Cover image for Long-Term Gamification: A Survey"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Long-Term Gamification: A Survey"}</h4>
              <p className={"conference"}>{"HCII | July 2024"}</p>
              <p className={"author"}>{"Lei Huang, Chao Deng, Jennifer Hoffman, Reza Hadi Mogavi, Justin Juho Kim, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1007/978-3-031-60692-2_3"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1007978-3-031-60695-3_11.png"} alt={"Cover image for Which Exergame is Better for Older Adults? An Exploratory Study on User Perspectives of Virtual Reality, Exercube, and 2D Exergames"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Which Exergame is Better for Older Adults? An Exploratory Study on User Perspectives of Virtual Reality, Exercube, and 2D Exergames"}</h4>
              <p className={"conference"}>{"HCII | July 2024"}</p>
              <p className={"author"}>{"Chao Deng, Jennifer Hoffman, Reza Hadi Mogavi, Juhyung Son, Simin Yang, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1007/978-3-031-60695-3_11"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/None.png"} alt={"Cover image for A Matter of Orientation: Interactive Artwork Recasting Historical Artifacts in Latent Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A Matter of Orientation: Interactive Artwork Recasting Historical Artifacts in Latent Reality"}</h4>
              <p className={"conference"}>{"ISEA | June 2024"}</p>
              <p className={"author"}>{"Zhiwan Cheung, Oksana Kryzhanivska, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/None"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/None.png"} alt={"Cover image for Pomotion: Generative Calligraphy Visualizing Emotion in Poems"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Pomotion: Generative Calligraphy Visualizing Emotion in Poems"}</h4>
              <p className={"conference"}>{"ISEA | June 2024"}</p>
              <p className={"author"}>{"Tiancheng Liu, Anqi Wang, Xinda Chen, Jing Yan, Yan Li, Pan Hui, and Kang Zhang"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/None"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1609icwsm.v18i1.31430.png"} alt={"Cover image for A Measurement Study of the Partisan News Sharing in the Russian Invasion of Ukraine in 2022"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A Measurement Study of the Partisan News Sharing in the Russian Invasion of Ukraine in 2022"}</h4>
              <p className={"conference"}>{"ICWSM | June 2024"}</p>
              <p className={"author"}>{"Yiming Zhu, Ehsan-Ul Haq, Gareth Tyson, Lik-Hang Lee, Yuyang Wang, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1609/icwsm.v18i1.31430"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453613904.3642105.png"} alt={"Cover image for CharacterMeet:... Supporting Creative Writers' Entire Story Character Construction Process Through Conversation with LLM-Powered Chatbot Avatars"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"CharacterMeet:... Supporting Creative Writers' Entire Story Character Construction Process Through Conversation with LLM-Powered Chatbot Avatars"}</h4>
              <p className={"conference"}>{"CHI | May 2024"}</p>
              <p className={"author"}>{"Hua Xuan Qin, Shan Jin, Mingming Fan, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3613904.3642105"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2403.14665.png"} alt={"Cover image for Sora OpenAI's Prelude: Social Media Perspectives on Sora OpenAI and the Future of AI Video Generation"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Sora OpenAI's Prelude: Social Media Perspectives on Sora OpenAI and the Future of AI Video Generation"}</h4>
              <p className={"conference"}>{"CHI Workshop | May 2024"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Derrick Wang, Joseph Tu, Hilda Hadan, Sabrina A. Sgandurra, Pan Hui, Lennart E. Nacke"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2403.14665"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2402.01697.png"} alt={"Cover image for APT-Pipe: An Automatic Prompt-Tuning Tool for Social Computing Data Annotation"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"APT-Pipe: An Automatic Prompt-Tuning Tool for Social Computing Data Annotation"}</h4>
              <p className={"conference"}>{"WWW | May 2024"}</p>
              <p className={"author"}>{"Yiming Zhu, Zhizhuo Yin, Gareth Tyson, Ehsan-ul Haq, Lik-hang Lee, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2402.01697"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453589334.3645334.png"} alt={"Cover image for Social Media Discourses on Interracial Intimacy: Tracking Racism and Sexism through Chinese Geo-located Social Media Data"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Social Media Discourses on Interracial Intimacy: Tracking Racism and Sexism through Chinese Geo-located Social Media Data"}</h4>
              <p className={"conference"}>{"WWW | May 2024"}</p>
              <p className={"author"}>{"Zheng Wei, Yixuan Xie, Danyun Xiao, Simin Zhang, Pan Hui, and Muzhi Zhou"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3589334.3645334"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453589335.3652003.png"} alt={"Cover image for Digital Democracy at Crossroads: A Meta-Analysis of Web and AI Influence on Global Elections"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Digital Democracy at Crossroads: A Meta-Analysis of Web and AI Influence on Global Elections"}</h4>
              <p className={"conference"}>{"WWW History Track | May 2024"}</p>
              <p className={"author"}>{"Zheng Wei, Xian Xu, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3589335.3652003"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453589335.3652000.png"} alt={"Cover image for History in Making: Political Campaigns in the Era of Artificial Intelligence-Generated Content"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"History in Making: Political Campaigns in the Era of Artificial Intelligence-Generated Content"}</h4>
              <p className={"conference"}>{"WWW History Track | May 2024"}</p>
              <p className={"author"}>{"Ehsan-ul Haq, Yiming Zhu, Pan Hui, and Gareth Tyson"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3589335.3652000"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2404.04693.png"} alt={"Cover image for OmniColor: A Global Camera Pose Optimization Approach of LiDAR-360Camera Fusion for Colorizing Point Clouds"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"OmniColor: A Global Camera Pose Optimization Approach of LiDAR-360Camera Fusion for Colorizing Point Clouds"}</h4>
              <p className={"conference"}>{"ICRA | May 2024"}</p>
              <p className={"author"}>{"Bonan Liu, Guoyang Zhao, Jianhao Jiao, Guang Cai, Chengyang Li,Handi Yin, Yuyang Wang, Ming Liu, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2404.04693"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453625468.3647612.png"} alt={"Cover image for FovOptix: Human Vision-Compatible Video Encoding and Adaptive Streaming in VR Cloud Gaming"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"FovOptix: Human Vision-Compatible Video Encoding and Adaptive Streaming in VR Cloud Gaming"}</h4>
              <p className={"conference"}>{"ACM MMSys | April 2024"}</p>
              <p className={"author"}>{"Ahmad Alhilal, Ze Wu, Yuk Hang Tsui, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3625468.3647612"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/ 10.1109VR58804.2024.00090.png"} alt={"Cover image for Text2VRScene:... Exploring the Paradigm of Automated Generation System for VR Experience From the Ground Up"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Text2VRScene:... Exploring the Paradigm of Automated Generation System for VR Experience From the Ground Up"}</h4>
              <p className={"conference"}>{"IEEE VR | March 2024"}</p>
              <p className={"author"}>{"Zhizhuo Yin, Yuyang Wang,Theodoros Papatheodorou, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/ 10.1109/VR58804.2024.00090"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109VR58804.2024.00029.png"} alt={"Cover image for Jump Cut Effects in Cinematic Virtual Reality: Editing with the 30-degree Rule and 180-degree Rule"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Jump Cut Effects in Cinematic Virtual Reality: Editing with the 30-degree Rule and 180-degree Rule"}</h4>
              <p className={"conference"}>{"IEEE VR | March 2024"}</p>
              <p className={"author"}>{"Junjie Zhang, Lik-Hang Lee, Yuyang Wang, Shan Jin, Danlu Fei, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/VR58804.2024.00029"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TVCG.2024.3372085.png"} alt={"Cover image for Dream360: Diverse and Immersive Outdoor Virtual Scene Creation via Transformer-Based 360? Image Outpainting"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Dream360: Diverse and Immersive Outdoor Virtual Scene Creation via Transformer-Based 360? Image Outpainting"}</h4>
              <p className={"conference"}>{"IEEE VR | March 2024"}</p>
              <p className={"author"}>{"Hao Ai, Zidong Cao, Haonan Lu, Chen Chen, Jian Ma, Pengyuan Zhou, Tae-Kyun Kim, Pan Hui, and Lin Wang"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TVCG.2024.3372085"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109PerCom59722.2024.10494441.png"} alt={"Cover image for AnchorLoc: Large-scale, Real-Time Visual Localization through Anchor Extraction and Detection"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"AnchorLoc: Large-scale, Real-Time Visual Localization through Anchor Extraction and Detection"}</h4>
              <p className={"conference"}>{"PerCom | March 2024"}</p>
              <p className={"author"}>{"Chun Ho Park, Ahmad Alhilal,Tristan Braud, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/PerCom59722.2024.10494441"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TWC.2024.3380820.png"} alt={"Cover image for Attention-based QoE-aware Digital Twin Empowered Edge Computing for Immersive Virtual Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Attention-based QoE-aware Digital Twin Empowered Edge Computing for Immersive Virtual Reality"}</h4>
              <p className={"conference"}>{"IEEE TWC | March 2024"}</p>
              <p className={"author"}>{"Jiadong Yu, Ahmad Alhilal, Tailin Zhou, Pan Hui, and Danny Tsang"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TWC.2024.3380820"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.108010447318.2024.2327199.png"} alt={"Cover image for Understanding Perceived User Reachability in Mobile UIs Using Data Analytics and Machine Learning"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Understanding Perceived User Reachability in Mobile UIs Using Data Analytics and Machine Learning"}</h4>
              <p className={"conference"}>{"IJHCI | March 2024"}</p>
              <p className={"author"}>{"Lik-Hang Lee, Yui-Pan Yau, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1080/10447318.2024.2327199"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
      </section>
      <section id={"yr2023"}>
        <br />
        <br />
        <h3>{"YEAR 2023"}</h3>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.2139ssrn.4665569.png"} alt={"Cover image for Legal Implications of Self-presence in the Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Legal Implications of Self-presence in the Metaverse"}</h4>
              <p className={"conference"}>{"Media & Arts Law Review | December 2023"}</p>
              <p className={"author"}>{"Jyh-An Lee, Liang Yang, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.2139/ssrn.4665569"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453639825.png"} alt={"Cover image for Behave Differently when Clustering: a Semi-Asynchronous Federated Learning Approach for IoT"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Behave Differently when Clustering: a Semi-Asynchronous Federated Learning Approach for IoT"}</h4>
              <p className={"conference"}>{"ACM Transactions on Sensor Networks | December 2023"}</p>
              <p className={"author"}>{"Boyu Fan, Xiang Su, Tarkoma Sasu, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3639825"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1007s10055-023-00934-5.png"} alt={"Cover image for Using a virtual reality interview simulator to explore factors influencing people's behavior"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Using a virtual reality interview simulator to explore factors influencing people's behavior"}</h4>
              <p className={"conference"}>{"Virtual Reality | December 2023"}</p>
              <p className={"author"}>{"Xinyi Luo, Yuyang Wang, Lik-Hang Lee, Zihan Xing, Shan Jin, Boya Dong, Yuanyi Hu, Zeming Chen, Jing Yan, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1007/s10055-023-00934-5"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109GLOBECOM54140.2023.10437100.png"} alt={"Cover image for QoE Optimization for VR Streaming: a Continual RL Framework in Digital Twin-empowered MEC"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"QoE Optimization for VR Streaming: a Continual RL Framework in Digital Twin-empowered MEC"}</h4>
              <p className={"conference"}>{"IEEE GLOBECOM | December 2023"}</p>
              <p className={"author"}>{"Jiadong Yu, Ahmad Alhilal, Tailin Zhou, Pan Hui, and Danny Tsang"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/GLOBECOM54140.2023.10437100"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453625007.3627595.png"} alt={"Cover image for Understanding Characteristics of Catalyst Users in the WallStreetBets Community"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Understanding Characteristics of Catalyst Users in the WallStreetBets Community"}</h4>
              <p className={"conference"}>{"IEEE/ACM ASONAM 23 | November 2023"}</p>
              <p className={"author"}>{"Ehsan-Ul Haq, Yiming Zhu, Zijun Lin, Haodi Weng, Gareth Tyson, Lik-Hang Lee, Reza Hadi Mogavi, Tristan Braud, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3625007.3627595"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453625007.3627475.png"} alt={"Cover image for Echo Chambers within the Russo-Ukrainian War: The Role of Bipartisan Users"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Echo Chambers within the Russo-Ukrainian War: The Role of Bipartisan Users"}</h4>
              <p className={"conference"}>{"IEEE/ACM ASONAM 23 | November 2023"}</p>
              <p className={"author"}>{"Peixian Zhang, Ehsan-Ul Haq, Yiming Zhu, Pan Hui, and Gareth Tyson"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3625007.3627475"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2306.11390.png"} alt={"Cover image for An Analysis of Twitter Discourse on the War Between Russia and Ukraine"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"An Analysis of Twitter Discourse on the War Between Russia and Ukraine"}</h4>
              <p className={"conference"}>{"IEEE/ACM ASONAM 23 | November 2023"}</p>
              <p className={"author"}>{"Haris Bin Zia, Ehsan-Ul Haq, Ignacio Castro, Pan Hui, and Gareth Tyson"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2306.11390"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1016j.wse.2023.12.002.png"} alt={"Cover image for Head-mounted display-based augmented reality for water quality visualisation"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Head-mounted display-based augmented reality for water quality visualisation"}</h4>
              <p className={"conference"}>{"Water Science and Engineering | November 2023"}</p>
              <p className={"author"}>{"Jacky Cao, Xiaoli Liu, Xiang Su, Jonas Eilertsen Hædahl, Thomas Berg Fjellestad, Donjete Haziri, André Hoang-An Vu, Jari Koskiaho, Satu Maaria Karjalainen, Anna-kaisa Ronkanen, Sasu Tarkoma, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1016/j.wse.2023.12.002"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453581783.3612438.png"} alt={"Cover image for Designing Loving-Kindness Meditation in Virtual Reality for Long-Distance Romantic Relationships"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Designing Loving-Kindness Meditation in Virtual Reality for Long-Distance Romantic Relationships"}</h4>
              <p className={"conference"}>{"ACM Multimedia 2023 | October 2023"}</p>
              <p className={"author"}>{"Xian Wang, Xiaoyu Mo, Lik-Hang Lee, Xiaoying Wei, Xiaofu Jin, Mingming Fan, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3581783.3612438"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453581783.3612580.png"} alt={"Cover image for Feeling Present! From Physical to Virtual Cinematography Lighting Education with Metashadow"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Feeling Present! From Physical to Virtual Cinematography Lighting Education with Metashadow"}</h4>
              <p className={"conference"}>{"ACM Multimedia 2023 | October 2023"}</p>
              <p className={"author"}>{"Zheng Wei, Xian Xu, Lik-Hang Lee, Wai Tong, Huamin Qu, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3581783.3612580"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453607546.3616806.png"} alt={"Cover image for Towards Optimising Transport Protocols on the 5G Edge for Mobile Augmented Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards Optimising Transport Protocols on the 5G Edge for Mobile Augmented Reality"}</h4>
              <p className={"conference"}>{"2nd International Workshop on Interactive eXtended Reality (iXR) | October 2023"}</p>
              <p className={"author"}>{"Jacky Cao, Xiang Su, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3607546.3616806"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TMC.2023.3329444.png"} alt={"Cover image for Towards Risk-averse Edge Computing with Deep Reinforcement Learning"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards Risk-averse Edge Computing with Deep Reinforcement Learning"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Mobile Computing | October 2023"}</p>
              <p className={"author"}>{"Dianlei Xu, Xiang Su, Huandong Wang, Sasu Tarkoma, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TMC.2023.3329444"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TCSS.2023.3330071.png"} alt={"Cover image for The Price is Right? The Economic Value of Sensor Sharing"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"The Price is Right? The Economic Value of Sensor Sharing"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Computational Social Systems | October 2023"}</p>
              <p className={"author"}>{"Ngoc Thi Nguyen, Maria Zubair, Agustin Zuniga, Sasu Tarkoma, Pan Hui, Hyowon Lee, Simon Tangi Perrault, Mostafa H. Ammar, Huber Flores, and Petteri Nurmi"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TCSS.2023.3330071"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1016j.chbah.2023.100027.png"} alt={"Cover image for ChatGPT in Education: A Blessing or a Curse? A Qualitative Study Exploring Early Adopters' Utilization and Perceptions"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"ChatGPT in Education: A Blessing or a Curse? A Qualitative Study Exploring Early Adopters' Utilization and Perceptions"}</h4>
              <p className={"conference"}>{"Computers in Human Behavior: Artificial Humans | October 2023"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Chao Deng, Justin Juho Kim, Pengyuan Zhou, Young D. Kwon, Ahmed Hosny Saleh Metwally, Ahmed Tlili, Simone Bassanelli, Antonio Bucchiarone, Sujit Gujar, Lennart E. Nacke, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1016/j.chbah.2023.100027"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2304.05984.png"} alt={"Cover image for A Deep Cybersickness Predictor through Kinematic Data with Encoded Physiological Representation"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A Deep Cybersickness Predictor through Kinematic Data with Encoded Physiological Representation"}</h4>
              <p className={"conference"}>{"22nd IEEE ISMAR | October 2023"}</p>
              <p className={"author"}>{"Ruichen Li, Yuyang Wang, Jean-Rémy Chardonnet, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2304.05984"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453623565.3623712.png"} alt={"Cover image for SoK: Distributed Computing in ICN"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"SoK: Distributed Computing in ICN"}</h4>
              <p className={"conference"}>{"10th ACM ICN 2023 | October 2023"}</p>
              <p className={"author"}>{"Wei Geng, Yulong Zhang, Dirk Kutscher, Abhishek Kumar, Sasu Tarkoma, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3623565.3623712"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453603163.3609028.png"} alt={"Cover image for Ghost Booking as a New Philanthropy Channel: A Case Study on Ukraine-Russia Conflict"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Ghost Booking as a New Philanthropy Channel: A Case Study on Ukraine-Russia Conflict"}</h4>
              <p className={"conference"}>{"34th ACM HT'23 | September 2023"}</p>
              <p className={"author"}>{"Fachrina Dewi Puspitasari, Gareth Tyson, Ehsan-Ul Haq, Pan Hui, and Lik-Hang Lee"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3603163.3609028"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2211.08700.png"} alt={"Cover image for Bi-directional Digital Twin and Edge Computing in the Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Bi-directional Digital Twin and Edge Computing in the Metaverse"}</h4>
              <p className={"conference"}>{"IEEE IoT Magazine | September 2023"}</p>
              <p className={"author"}>{"Jiadong Yu, Ahmad Alhilal, Pan Hui, and Danny Tsang"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2211.08700"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109MCOMSTD.0008.2200067.png"} alt={"Cover image for Towards a Traffic Metaverse with Shared Vehicle Perception"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards a Traffic Metaverse with Shared Vehicle Perception"}</h4>
              <p className={"conference"}>{"IEEE Communications Standards Magazine | September 2023"}</p>
              <p className={"author"}>{"Ahmad Alhilal, Tristan Braud, Lik-Hang Lee, Hang Chen, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/MCOMSTD.0008.2200067"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453615522.3615535.png"} alt={"Cover image for Design and Implementation of a Virtual Reality Simulator for Enhancing Chemistry Education"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Design and Implementation of a Virtual Reality Simulator for Enhancing Chemistry Education"}</h4>
              <p className={"conference"}>{"16th VINCI 2023 | September 2023"}</p>
              <p className={"author"}>{"Shan Jin, Yuyang Wang, Lik-Hang Lee, Xinyi Luo, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3615522.3615535"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453580305.3599853.png"} alt={"Cover image for Large-scale Urban Cellular Traffic Generation via Knowledge-Enhanced GANs with Multi-Periodic Patterns"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Large-scale Urban Cellular Traffic Generation via Knowledge-Enhanced GANs with Multi-Periodic Patterns"}</h4>
              <p className={"conference"}>{"29th ACM SIGKDD | August 2023"}</p>
              <p className={"author"}>{"Shuodi Hui, Huandong Wang, Tong Li, Xinghao Yang, Junlan Feng, Lin Zhu, Chao Deng, Pan Hui, Depeng Jin, and Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3580305.3599853"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2308.00465.png"} alt={"Cover image for A Satellite Imagery Dataset for Long-Term Sustainable Development in United States Cities"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A Satellite Imagery Dataset for Long-Term Sustainable Development in United States Cities"}</h4>
              <p className={"conference"}>{"Nature Scientific Data | July 2023"}</p>
              <p className={"author"}>{"Yanxin Xi, Yu Liu, Tong Li, Jintao Ding, Yunke Zhang, Sasu Tarkoma, Yong Li, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2308.00465"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2302.05623.png"} alt={"Cover image for Your Favorite Gameplay Speaks Volumes about You: Predicting User Behavior and Hexad Type"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Your Favorite Gameplay Speaks Volumes about You: Predicting User Behavior and Hexad Type"}</h4>
              <p className={"conference"}>{"25th HCII 2023 | July 2023"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Chao Deng, Jennifer Hoffman, Ehsan-Ul Haq, Sujit Gujar, Antonio Bucchiarone, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2302.05623"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453573051.3596185.png"} alt={"Cover image for Envisioning an Inclusive Metaverse: Student Perspectives on Accessible and Empowering Metaverse-Enabled Learning"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Envisioning an Inclusive Metaverse: Student Perspectives on Accessible and Empowering Metaverse-Enabled Learning"}</h4>
              <p className={"conference"}>{"10th ACM L@S 2023 | July 2023"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Jennifer Hoffman, Chao Deng, Yiwei Du, Ehsan-Ul Haq, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3573051.3596185"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ICDCSW60045.2023.00022.png"} alt={"Cover image for Empowering the Metaverse with Generative AI: Survey and Future Directions"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Empowering the Metaverse with Generative AI: Survey and Future Directions"}</h4>
              <p className={"conference"}>{"IEEE SocialMeta 2023 | July 2023"}</p>
              <p className={"author"}>{"Hua Xuan Qin, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ICDCSW60045.2023.00022"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ICDCSW60045.2023.00025.png"} alt={"Cover image for Network Traffic in the Metaverse: The Case of Social VR"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Network Traffic in the Metaverse: The Case of Social VR"}</h4>
              <p className={"conference"}>{"IEEE SocialMeta 2023 | July 2023"}</p>
              <p className={"author"}>{"Ahmad Alhilal, Kirill Shatilov, Gareth Tyson, Tristan Braud, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ICDCSW60045.2023.00025"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1609icwsm.v17i1.22132.png"} alt={"Cover image for Getting Back on Track: Understanding COVID-19 Impact on Urban Mobility and Segregation with Location Service Data"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Getting Back on Track: Understanding COVID-19 Impact on Urban Mobility and Segregation with Location Service Data"}</h4>
              <p className={"conference"}>{"17th ICWSM | June 2023"}</p>
              <p className={"author"}>{"Lin Chen, Fengli Xu, Qianyue Hao, Yong Li, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1609/icwsm.v17i1.22132"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453597063.3597360.png"} alt={"Cover image for Players are not Ready 101: A Tutorial on Organising Mixed-mode Events in the Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Players are not Ready 101: A Tutorial on Organising Mixed-mode Events in the Metaverse"}</h4>
              <p className={"conference"}>{"1st ACM Workshop on Metaverse Systems and Applications | June 2023"}</p>
              <p className={"author"}>{"Kirill Shatilov, Ahmad Alhilal, Tristan Braud, Lik-hang Lee, Pengyuan Zhou, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3597063.3597360"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453591129.png"} alt={"Cover image for Exploring Gaze-assisted and Hand-based Region Selection in Augmented Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Exploring Gaze-assisted and Hand-based Region Selection in Augmented Reality"}</h4>
              <p className={"conference"}>{"PACM HCI | May 2023"}</p>
              <p className={"author"}>{"Rongkai Shi, Yushi Wei, Xueying Qin, Pan Hui, and Hai-Ning Liang"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3591129"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/ICLT_2023.png"} alt={"Cover image for Metaverse Enhanced Project-based Learning: Experiences from An Interdisciplinary University"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Metaverse Enhanced Project-based Learning: Experiences from An Interdisciplinary University"}</h4>
              <p className={"conference"}>{"ICLT | May 2023"}</p>
              <p className={"author"}>{"Qingqing Xing, Yuyang Wang, Jiayang Huang, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/-"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453543507.3587432.png"} alt={"Cover image for Tangible Web: An Interactive Immersion Virtual Reality System That Travels Across Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Tangible Web: An Interactive Immersion Virtual Reality System That Travels Across Reality"}</h4>
              <p className={"conference"}>{"32nd ACM WWW | April/May 2023"}</p>
              <p className={"author"}>{"Simin Yang, Reza Hadi Mogavi, Tristan Braud, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3543507.3587432"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2305.00510.png"} alt={"Cover image for Towards Computational Architecture of Liberty: A Comprehensive Survey on Deep Learning for Generating Virtual Architecture in the Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards Computational Architecture of Liberty: A Comprehensive Survey on Deep Learning for Generating Virtual Architecture in the Metaverse"}</h4>
              <p className={"conference"}>{"Technical Report | April 2023"}</p>
              <p className={"author"}>{"Anqi Wang, Jiahua Dong, Jiachuan Shen, Lik-Hang Lee, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2305.00510"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453589344.png"} alt={"Cover image for Learning Representations of Satellite Imagery by Leveraging Point-of-Interests"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Learning Representations of Satellite Imagery by Leveraging Point-of-Interests"}</h4>
              <p className={"conference"}>{"ACM TIST | March 2023"}</p>
              <p className={"author"}>{"Tong Li, Yanxin Xi, Huandong Wang, Yong Li, Sasu Tarkoma, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3589344"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TMC.2023.3258974.png"} alt={"Cover image for Learn to Optimize the Constrained Shortest Path on Large Dynamic Graphs"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Learn to Optimize the Constrained Shortest Path on Large Dynamic Graphs"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Mobile Computing | March 2023"}</p>
              <p className={"author"}>{"Jiaming Yin, Weixiong Rao, Qinpei Zhao, Chenxi Zhang, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TMC.2023.3258974"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.108010447318.2023.2188799.png"} alt={"Cover image for The Dark Side of Augmented Reality: Exploring Manipulative Designs in AR"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"The Dark Side of Augmented Reality: Exploring Manipulative Designs in AR"}</h4>
              <p className={"conference"}>{"International Journal of Human-Computer Interaction | March 2023"}</p>
              <p className={"author"}>{"Xian Wang, Lik-Hang Lee, Carlos Bermejo Fernandez, and Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1080/10447318.2023.2188799"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1016j.dib.2023.108898.png"} alt={"Cover image for A Human Mobility Dataset Collected via LBSLab"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A Human Mobility Dataset Collected via LBSLab"}</h4>
              <p className={"conference"}>{"Elsevier Data in Brief | February 2023"}</p>
              <p className={"author"}>{"Yuwei Zhang, Qingyuan Gong, Yang Chen, Yu Xiao, Xin Wang, Pan Hui, and Xiaoming Fu"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1016/j.dib.2023.108898"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TLT.2023.3255882.png"} alt={"Cover image for Can Underprivileged Children Learn Effectively at Home? A Six-Month Study of Game-based Traditional Chinese Learning During the Pandemic Lockdown"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Can Underprivileged Children Learn Effectively at Home? A Six-Month Study of Game-based Traditional Chinese Learning During the Pandemic Lockdown"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Learning Technologies | February 2023"}</p>
              <p className={"author"}>{"Ka-Yan Fung, Lik-Hang Lee, Pan Hui, and Shenghui Song"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TLT.2023.3255882"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
      </section>
      <section id={"yr2022"}>
        <br />
        <br />
        <h3>{"YEAR 2022"}</h3>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453357384.3357971.png"} alt={"Cover image for Accounts in Online Developer Communities Using Deep Learning"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Accounts in Online Developer Communities Using Deep Learning"}</h4>
              <p className={"conference"}>{"IEEE TKDE | December 2022"}</p>
              <p className={"author"}>{"Qingyuan Gong, Yushan Liu, Jiayun Zhang, Yang Chen, Qi Li, Yu Xiao, Xin Wang, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3357384.3357971"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453551626.3564936.png"} alt={"Cover image for Human-Avatar Interaction in Metaverse: Framework for Full-body Interaction"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Human-Avatar Interaction in Metaverse: Framework for Full-body Interaction"}</h4>
              <p className={"conference"}>{"ACM Multimedia Asia | December 2022"}</p>
              <p className={"author"}>{"Kit-Yung Lam, Ahmad Alhilal, Liang Yang, Lik-Hang Lee, Gareth Tyson, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3551626.3564936"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109PerCom53586.2022.9762389.png"} alt={"Cover image for Context-driven Encrypted Multimedia Traffic Classification on Mobile Devices"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Context-driven Encrypted Multimedia Traffic Classification on Mobile Devices"}</h4>
              <p className={"conference"}>{"Elsevier Pervasive and Mobile Computing | December 2022"}</p>
              <p className={"author"}>{"Mohammad Hoque, Benjamin Finley, Ashwin Rao, Abhishek Kumar, Pan Hui, Mostafa Ammar, Sasu Tarkoma"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/PerCom53586.2022.9762389"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1016j.simpa.2022.100436.png"} alt={"Cover image for Federated Split GANs for Collaborative Training with Heterogeneous Devices"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Federated Split GANs for Collaborative Training with Heterogeneous Devices"}</h4>
              <p className={"conference"}>{"Elsevier Software Impacts | December 2022"}</p>
              <p className={"author"}>{"Yilei Liang, Pranvera Kortoçi, Pengyuan Zhou, Lik-Hang Lee, Abbas Mehrabi, Pan Hui, Sasu Tarkoma, Jon Crowcroft"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1016/j.simpa.2022.100436"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109HPCC-DSS-SmartCity-DependSys57074.2022.00090.png"} alt={"Cover image for Fairness-Aware Algorithms for Seed Allocation in Social Advertising"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Fairness-Aware Algorithms for Seed Allocation in Social Advertising"}</h4>
              <p className={"conference"}>{"IEEE HPCC | December 2022"}</p>
              <p className={"author"}>{"Pengzi Wang, Yiming Zhu, Kai Han, Zhizhuo Yin, Qing Xiu, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/HPCC-DSS-SmartCity-DependSys57074.2022.00090"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453555636.png"} alt={"Cover image for It's all Relative! A Method to Counter Human Bias in Crowdsourced Stance Detection of News Articles"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"It's all Relative! A Method to Counter Human Bias in Crowdsourced Stance Detection of News Articles"}</h4>
              <p className={"conference"}>{"ACM CSCW | November 2022"}</p>
              <p className={"author"}>{"Ehsan-Ul Haq, Yang K. Lu, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3555636"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453555553.png"} alt={"Cover image for More Gamification Is Not Always Better: A Case Study of Promotional Gamification in a Question Answering Website"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"More Gamification Is Not Always Better: A Case Study of Promotional Gamification in a Question Answering Website"}</h4>
              <p className={"conference"}>{"ACM CSCW | November 2022"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Ehsan-Ul Haq, Sujit Gujar, Pan Hui, Xiaojuan Ma"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3555553"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453555124.png"} alt={"Cover image for What Do Users Think of Promotional Gamification Schemes? A Qualitative Case Study in a Question Answering Website"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"What Do Users Think of Promotional Gamification Schemes? A Qualitative Case Study in a Question Answering Website"}</h4>
              <p className={"conference"}>{"ACM CSCW | November 2022"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Yuanhao Zhang, Ehsan-Ul Haq, Yongjin Wu, Pan Hui, Xiaojuan Ma"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3555124"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2210.13582.png"} alt={"Cover image for A causal Analysis on the Anchor Store Effect Using Open-LBSN Data"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A causal Analysis on the Anchor Store Effect Using Open-LBSN Data"}</h4>
              <p className={"conference"}>{"IEEE/ACM ASONAM | November 2022"}</p>
              <p className={"author"}>{"Anish K. Vallapuram, YoungDae Kwon, Lik-Hang Lee, Fengli Xu, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2210.13582"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ASONAM55673.2022.10068611.png"} alt={"Cover image for Psychologists, Therapists, Writers, Doctors? Exploring Mental Health Communications among Instagram Coaches"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Psychologists, Therapists, Writers, Doctors? Exploring Mental Health Communications among Instagram Coaches"}</h4>
              <p className={"conference"}>{"IEEE/ACM ASONAM | November 2022"}</p>
              <p className={"author"}>{"Ehsan Ul Haq, Lik-Hang Lee, Gareth Tyson, Reza Hadi Mogavi, Tristan Braud, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ASONAM55673.2022.10068611"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453536221.3558176.png"} alt={"Cover image for Decentralized, not Dehumanized: Bringing Utility to NFTs through Multimodal Interaction"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Decentralized, not Dehumanized: Bringing Utility to NFTs through Multimodal Interaction"}</h4>
              <p className={"conference"}>{"ACM ICMI | November 2022"}</p>
              <p className={"author"}>{"Anqi Wang, Lik-Hang Lee, Tristan Braud, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3536221.3558176"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453536221.3558174.png"} alt={"Cover image for Beyond the Blue Sky of Multimodal Interaction: A Centennial Vision of Interplanetary Virtual Spaces in Turn-based Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Beyond the Blue Sky of Multimodal Interaction: A Centennial Vision of Interplanetary Virtual Spaces in Turn-based Metaverse"}</h4>
              <p className={"conference"}>{"ACM ICMI | November 2022"}</p>
              <p className={"author"}>{"Lik Hang Lee, Carlos Bermejo Fernandez, Ahmad Alhilal, Tristan Braud, Simo Hosio, Esmee De Haas, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3536221.3558174"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453565698.3565781.png"} alt={"Cover image for Reducing Stress and Anxiety in the Metaverse: A Systematic Review of Meditation, Mindfulness and Virtual Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Reducing Stress and Anxiety in the Metaverse: A Systematic Review of Meditation, Mindfulness and Virtual Reality"}</h4>
              <p className={"conference"}>{"Chinese CHI | November 2022"}</p>
              <p className={"author"}>{"Xian Wang, Xiaoyu Mo, Mingming Fan, Lik-Hang Lee, Bertram Shi, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3565698.3565781"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TMM.2022.3217397.png"} alt={"Cover image for Free3Net: Gliding Free, Orientation Free, and Anchor Free Network for Oriented Object Detection"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Free3Net: Gliding Free, Orientation Free, and Anchor Free Network for Oriented Object Detection"}</h4>
              <p className={"conference"}>{"IEEE TMM | October 2022"}</p>
              <p className={"author"}>{"Zhonghong Ou, Zhongjie Chen, Lina Fan, Meina Song, Zheng Yan, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TMM.2022.3217397"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109JIOT.2022.3215469.png"} alt={"Cover image for AD-RCNN: Adaptive Dynamic Neural Network For Small Object Detection"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"AD-RCNN: Adaptive Dynamic Neural Network For Small Object Detection"}</h4>
              <p className={"conference"}>{"IEEE IoT Journal | October 2022"}</p>
              <p className={"author"}>{"Zhonghong Ou, Zhaofengnian Wang, Fenrui Xiao, Baiqiao Xiong, Hongxing Zhang, Meina Song, Zheng Yan, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/JIOT.2022.3215469"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2210.06134.png"} alt={"Cover image for Identity, Crimes, and Law Enforcement in the Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Identity, Crimes, and Law Enforcement in the Metaverse"}</h4>
              <p className={"conference"}>{"90th INTERPOL General Assembly | October 2022"}</p>
              <p className={"author"}>{"Hua Xuan Qin, Yuyang Wang, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2210.06134"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453503161.3548252.png"} alt={"Cover image for PassWalk: Spatial Authentication Leveraging Lateral Shift and Gaze on Mobile Headsets"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"PassWalk: Spatial Authentication Leveraging Lateral Shift and Gaze on Mobile Headsets"}</h4>
              <p className={"conference"}>{"ACM MM | October 2022"}</p>
              <p className={"author"}>{"Abhishek Kumar, Lik-Hang Lee, Jagmohan Chauhan, Xiang Su, Mohammad A. Hoque, Susanna Pirttikangas, Sasu Tarkoma, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3503161.3548252"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109IROS47612.2022.9981123.png"} alt={"Cover image for Towards Reproducible Evaluations for Flying Drone Controllers in Virtual Environments"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards Reproducible Evaluations for Flying Drone Controllers in Virtual Environments"}</h4>
              <p className={"conference"}>{"IEEE/RSJ IROS | October 2022"}</p>
              <p className={"author"}>{"Zheng Li, Yiming Huang, Yui-Pan Yau, Pan Hui, Lik-Hang Lee"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/IROS47612.2022.9981123"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453511808.3557153.png"} alt={"Cover image for Predicting Multi-level Socioeconomic Indicators from Structural Urban Imagery"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Predicting Multi-level Socioeconomic Indicators from Structural Urban Imagery"}</h4>
              <p className={"conference"}>{"ACM CIKM | October 2022"}</p>
              <p className={"author"}>{"Tong Li, Shiduo Xin, Yanxin Xi, Sasu Tarkoma, Pan Hui, Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3511808.3557153"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ACCESS.2022.3210985.png"} alt={"Cover image for Street Smart in 5G: Vehicular Applications, Communication"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Street Smart in 5G: Vehicular Applications, Communication"}</h4>
              <p className={"conference"}>{"IEEE Access | September 2022"}</p>
              <p className={"author"}>{"Ahmad Alhilal, Tristan Braud, Benjamin Finley, Dongzhe Su, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ACCESS.2022.3210985"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109MMUL.2022.3211351.png"} alt={"Cover image for DiOS - An Extended Reality Operating System for the Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"DiOS - An Extended Reality Operating System for the Metaverse"}</h4>
              <p className={"conference"}>{"IEEE MultiMedia | September 2022"}</p>
              <p className={"author"}>{"Tristan Braud, Lik-Hang Lee, Ahmad Alhilal, Carlos Bermejo, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/MMUL.2022.3211351"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453556557.3557953.png"} alt={"Cover image for Federated Split GANs"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Federated Split GANs"}</h4>
              <p className={"conference"}>{"ACM FedEdge | October 2022"}</p>
              <p className={"author"}>{"Pranvera Kortoçi, Yilei Liang, Pengyuan Zhou, Lik-Hang Lee, Abbas Mehrabi, Pan Hui, Sasu Tarkoma, Jon Crowcroft"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org//10.1145/3556557.3557953"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453534585.png"} alt={"Cover image for Are You Left Out? An Efficient and Fair Federated Learning for Personalized Profiles on Wearable Devices of Inferior Networking Conditions"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Are You Left Out? An Efficient and Fair Federated Learning for Personalized Profiles on Wearable Devices of Inferior Networking Conditions"}</h4>
              <p className={"conference"}>{"ACM IMWUT/UbiComp | September 2022"}</p>
              <p className={"author"}>{"Pengyuan Zhou, Hengwei Xu, Lik-Hang Lee, Pei Fang, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org//10.1145/3534585"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TMC.2022.3200401.png"} alt={"Cover image for Video Content Placement at the Network Edge: Centralized and Distributed Algorithms"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Video Content Placement at the Network Edge: Centralized and Distributed Algorithms"}</h4>
              <p className={"conference"}>{"IEEE TMC | August 2022"}</p>
              <p className={"author"}>{"Yanan Gao, Song Yang, Fan Li, Stojan Trajanovski, Pan Zhou, Pan Hui, Xiaoming Fu"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TMC.2022.3200401"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453557999.png"} alt={"Cover image for Mobile Augmented Reality: User Interfaces, Frameworks, and Intelligence"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Mobile Augmented Reality: User Interfaces, Frameworks, and Intelligence"}</h4>
              <p className={"conference"}>{"ACM CSUR | August 2022"}</p>
              <p className={"author"}>{"Jacky Cao, Kit-Yung Lam, Lik-Hang Lee, Xiaoli Liu, Pan Hui, Xiang Su"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3557999"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453552436.png"} alt={"Cover image for Hierarchical Multi-agent Model for Reinforced Medical Resource Allocation with Imperfect Information"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Hierarchical Multi-agent Model for Reinforced Medical Resource Allocation with Imperfect Information"}</h4>
              <p className={"conference"}>{"ACM TIST | July 2022"}</p>
              <p className={"author"}>{"Qianyue Hao, Fengli Xu, Lin Chen, Pan Hui, Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3552436"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TKDE.2022.3193128.png"} alt={"Cover image for CityNeuro: Towards Location and Time Prediction for Urban Abnormal Events"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"CityNeuro: Towards Location and Time Prediction for Urban Abnormal Events"}</h4>
              <p className={"conference"}>{"IEEE TKDE | July 2022"}</p>
              <p className={"author"}>{"Mingyang Zhang, Tong Li, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TKDE.2022.3193128"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1038s41562-022-01429-0.png"} alt={"Cover image for Strategic COVID-19 Vaccine Distribution Can Simultaneously Elevate Social Utility and Equity"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Strategic COVID-19 Vaccine Distribution Can Simultaneously Elevate Social Utility and Equity"}</h4>
              <p className={"conference"}>{"Nature Human Behaviour | July 2022"}</p>
              <p className={"author"}>{"Lin Chen, Fengli Xu, Zhenyu Han, Kun Tang, Pan Hui, James Evans, Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1038/s41562-022-01429-0"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ICDCSW56584.2022.00053.png"} alt={"Cover image for Re-shaping Post-COVID-19 Teaching and Learning: A Blueprint of Virtual-Physical Blended Classrooms in the Metaverse Era"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Re-shaping Post-COVID-19 Teaching and Learning: A Blueprint of Virtual-Physical Blended Classrooms in the Metaverse Era"}</h4>
              <p className={"conference"}>{"IEEE ICDCS Workshop | July 2022"}</p>
              <p className={"author"}>{"Yuyang Wang, Lik-Hang Lee, Tristan Braud, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ICDCSW56584.2022.00053"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ICDCSW56584.2022.00058.png"} alt={"Cover image for Life, the Metaverse and Everything: An Overview of Privacy, Ethics, and Governance in Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Life, the Metaverse and Everything: An Overview of Privacy, Ethics, and Governance in Metaverse"}</h4>
              <p className={"conference"}>{"IEEE ICDCS Workshop | July 2022"}</p>
              <p className={"author"}>{"Carlos Bermejo, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ICDCSW56584.2022.00058"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453511095.3536372.png"} alt={"Cover image for Weaponising Social Media for Information Divide and Warfare"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Weaponising Social Media for Information Divide and Warfare"}</h4>
              <p className={"conference"}>{"ACM HT | June 2022"}</p>
              <p className={"author"}>{"Ehsan Ul Haq, Gareth Tyson, Tristan Braud, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org//10.1145/3511095.3536372"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453524273.3528180.png"} alt={"Cover image for 3DeformR: Freehand 3D Model Editing in Virtual Environments Considering Head Movements on Mobile Headsets"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"3DeformR: Freehand 3D Model Editing in Virtual Environments Considering Head Movements on Mobile Headsets"}</h4>
              <p className={"conference"}>{"ACM MMSys | June 2022"}</p>
              <p className={"author"}>{"Kit-Yung Lam, Lik-Hang Lee, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3524273.3528180"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2111.05173.png"} alt={"Cover image for EdgeXAR: A 6-DoF Camera Multi-target Interaction Framework for MAR with User-friendly Latency Compensation"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"EdgeXAR: A 6-DoF Camera Multi-target Interaction Framework for MAR with User-friendly Latency Compensation"}</h4>
              <p className={"conference"}>{"ACM EICS | June 2022"}</p>
              <p className={"author"}>{"Wenxiao Zhang, Sikun Lin, Farshid Hassani, Haofei Cheng, Tristan Braud, Pengyuan Zhou, Lik-Hang Lee, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2111.05173"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453491140.3528274.png"} alt={"Cover image for When Gamification Spoils Your Learning: A Qualitative Case Study of Gamification Misuse in a Language-Learning App"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"When Gamification Spoils Your Learning: A Qualitative Case Study of Gamification Misuse in a Language-Learning App"}</h4>
              <p className={"conference"}>{"ACM L@S | June 2022"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Bingcan Guo, Yuanhao Zhang, Ehsan Ul Haq, Pan Hui, Xiaojuan Ma"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3491140.3528274"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453543434.3543642.png"} alt={"Cover image for Tips, Tidings, and Tech: Governmental Communication on Facebook During the COVID-19 Pandemic"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Tips, Tidings, and Tech: Governmental Communication on Facebook During the COVID-19 Pandemic"}</h4>
              <p className={"conference"}>{"DG.O | June 2022"}</p>
              <p className={"author"}>{"Ehsan Ul Haq, Tristan Braud, Lik Hang Lee, Reza Hadi Mogavi, He Zhang, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3543434.3543642"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ICASSP43922.2022.9747497.png"} alt={"Cover image for FedClean: A Defense Mechanism Against Parameter Poisoning Attacks in Federated Learning"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"FedClean: A Defense Mechanism Against Parameter Poisoning Attacks in Federated Learning"}</h4>
              <p className={"conference"}>{"IEEE ICASSP | May 2022"}</p>
              <p className={"author"}>{"Abhishek Kumar, Vivek Khimani, Dimitris Chatzopoulos, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ICASSP43922.2022.9747497"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2103.12542.png"} alt={"Cover image for EmgAuth: Unlocking Smartphones with EMG Signals"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"EmgAuth: Unlocking Smartphones with EMG Signals"}</h4>
              <p className={"conference"}>{"IEEE TMC | May 2022"}</p>
              <p className={"author"}>{"Boyu Fan, Xuefeng Liu, Xiang Su, Jianwei Niu, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2103.12542"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453485447.3512276.png"} alt={"Cover image for Nebula: Reliable Low-latency Video Transmission for Mobile Cloud Gaming"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Nebula: Reliable Low-latency Video Transmission for Mobile Cloud Gaming"}</h4>
              <p className={"conference"}>{"ACM WWW | April 2022"}</p>
              <p className={"author"}>{"Ahmad Alhilal, Tristan Braud, Bo Han, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3485447.3512276"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453485447.3512149.png"} alt={"Cover image for Beyond the First Law of Geography: Learning Representations of Satellite Imagery by Leveraging Point-of-Interests"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Beyond the First Law of Geography: Learning Representations of Satellite Imagery by Leveraging Point-of-Interests"}</h4>
              <p className={"conference"}>{"ACM WWW | April 2022"}</p>
              <p className={"author"}>{"Yanxin Xi, Tong Li, Huandong Wang, Yong Li, Sasu Tarkoma, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3485447.3512149"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453485447.3512268.png"} alt={"Cover image for Screenshots, Symbols, and Personal Thoughts: The Role of Instagram for Social Activism"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Screenshots, Symbols, and Personal Thoughts: The Role of Instagram for Social Activism"}</h4>
              <p className={"conference"}>{"ACM WWW | April 2022"}</p>
              <p className={"author"}>{"Ehsan Ul Haq, Tristan Braud, Yui-Pan Yau, Lik Hang Lee, Franziska Keller, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3485447.3512268"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453487553.3524202.png"} alt={"Cover image for Short, Colorful, and Irreverent! A Comparative Analysis of New Users on WallstreetBets During the Gamestop Short-squeeze"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Short, Colorful, and Irreverent! A Comparative Analysis of New Users on WallstreetBets During the Gamestop Short-squeeze"}</h4>
              <p className={"conference"}>{"ACM WWW Industry Track | April 2022"}</p>
              <p className={"author"}>{"Ehsan Ul Haq, Tristan Braud, Lik Hang Lee, Anish Krishna Vallapuram, Yue Yu, Gareth Tyson, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3487553.3524202"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109COMST.2022.3163176.png"} alt={"Cover image for Smartphone App Usage Analysis: Datasets, Methods, and Applications"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Smartphone App Usage Analysis: Datasets, Methods, and Applications"}</h4>
              <p className={"conference"}>{"IEEE COMST | March 2022"}</p>
              <p className={"author"}>{"Tong Li, Tong Xia, Huandong Wang, Zhen Tu, Sasu Tarkoma, Zhu Han, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/COMST.2022.3163176"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TVCG.2022.3150467.png"} alt={"Cover image for SEAR: Scaling Experiences in Multi-user Augmented Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"SEAR: Scaling Experiences in Multi-user Augmented Reality"}</h4>
              <p className={"conference"}>{"IEEE VR | March 2022"}</p>
              <p className={"author"}>{"Wenxiao Zhang, Bo Han, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TVCG.2022.3150467"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109VRW55335.2022.00044.png"} alt={"Cover image for Scaling-up AR: University Campus as a Physical-Digital Metaverse"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Scaling-up AR: University Campus as a Physical-Digital Metaverse"}</h4>
              <p className={"conference"}>{"IEEE VR Workshop | March 2022"}</p>
              <p className={"author"}>{"Tristan Braud, Carlos Bermejo, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/VRW55335.2022.00044"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1016j.pmcj.2022.101737.png"} alt={"Cover image for Context-driven Encrypted Multimedia Traffic Classification on Mobile Devices"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Context-driven Encrypted Multimedia Traffic Classification on Mobile Devices"}</h4>
              <p className={"conference"}>{"IEEE PerCom | March 2022"}</p>
              <p className={"author"}>{"Mohammad Hoque, Benjamin Finley, Ashwin Rao, Abhishek Kumar, Pan Hui, Mostafa Ammar, Sasu Tarkoma"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1016/j.pmcj.2022.101737"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109HAPTICS52432.2022.9765609.png"} alt={"Cover image for VibroWeight: Simulating Weight and Center of Gravity Changes of Objects in Virtual Reality for Enhanced Realism"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"VibroWeight: Simulating Weight and Center of Gravity Changes of Objects in Virtual Reality for Enhanced Realism"}</h4>
              <p className={"conference"}>{"IEEE HAPTICS | March 2022"}</p>
              <p className={"author"}>{"Xian Wang, Diego Monteiro, Lik-Hang Lee, Pan Hui, Hai-Ning Liang"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/HAPTICS52432.2022.9765609"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453508396.3512880.png"} alt={"Cover image for Implementing GDPR for Mobile and Ubiquitous Computing"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Implementing GDPR for Mobile and Ubiquitous Computing"}</h4>
              <p className={"conference"}>{"ACM HotMobile | March 2022"}</p>
              <p className={"author"}>{"Carlos Bermejo, Tristan Braud, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3508396.3512880"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2101.05508.png"} alt={"Cover image for Augmented Informative Cooperative Perception"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Augmented Informative Cooperative Perception"}</h4>
              <p className={"conference"}>{"IEEE TITS | February 2022"}</p>
              <p className={"author"}>{"Pengyuan Zhou, Pranvera Kortoci, Yui-Pan Yau, Tristan Braud, Xiujun Wang, Benjamin Finley, Lik-Hang Lee, Sasu Tarkoma, Jussi Kangasharju, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2101.05508"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TMC.2022.3156939.png"} alt={"Cover image for MyoBoard: Inertial Motion Sensing and Gesture-based QWERTY Keyboard for Extended Realities"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"MyoBoard: Inertial Motion Sensing and Gesture-based QWERTY Keyboard for Extended Realities"}</h4>
              <p className={"conference"}>{"IEEE TMC | February 2022"}</p>
              <p className={"author"}>{"Kirill A. Shatilov, Youngdae Kwon, Lik-Hang Lee, Dimitris Chatzopoulos, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TMC.2022.3156939"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109MPRV.2022.3152926.png"} alt={"Cover image for Toward City-Scale Litter Monitoring using Autonomous Ground Vehicles"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Toward City-Scale Litter Monitoring using Autonomous Ground Vehicles"}</h4>
              <p className={"conference"}>{"IEEE Pervasive Computing Magazine | February 2022"}</p>
              <p className={"author"}>{"Zhigang Yin, Mayowa Olapade, Mohan Liyanage, Farooq Dar, Agustin Zuniga, Naser Hossein Motlagh, Xiang Su, Sasu Tarkoma, Pan Hui, Petteri Nurmi, Huber Flores"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/MPRV.2022.3152926"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
      </section>
      <section id={"yr2021"}>
        <br />
        <br />
        <h3>{"YEAR 2021"}</h3>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2011.04218.png"} alt={"Cover image for Automorphic Equivalence-aware Graph Neural Network."} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Automorphic Equivalence-aware Graph Neural Network."}</h4>
              <p className={"conference"}>{"NeurIPS | November 2021"}</p>
              <p className={"author"}>{"Fengli Xu, Quanming Yao, Pan Hui, and Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2011.04218"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453485983.3494848.png"} alt={"Cover image for Talaria: In-engine Synchronisation for Seamless Migration of Mobile Edge Gaming Instances"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Talaria: In-engine Synchronisation for Seamless Migration of Mobile Edge Gaming Instances"}</h4>
              <p className={"conference"}>{"CoNEXT | December 2021"}</p>
              <p className={"author"}>{"Tristan Braud, Ahmad Alhilal, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3485983.3494848"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ICDM51629.2021.00191.png"} alt={"Cover image for Adaptive Spatio-Temporal Convolutional Network for Traffic Prediction"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Adaptive Spatio-Temporal Convolutional Network for Traffic Prediction"}</h4>
              <p className={"conference"}>{"ICDM | December 2021"}</p>
              <p className={"author"}>{"Mingyang Zhang, Yong Li, Funing Sun, Diansheng Guo, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ICDM51629.2021.00191"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2110.13290.png"} alt={"Cover image for Exploring System Performance of Continual Learning for Mobile and Embedded Sensing Applications"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Exploring System Performance of Continual Learning for Mobile and Embedded Sensing Applications"}</h4>
              <p className={"conference"}>{"SEC (Best Paper Award) | December 2021"}</p>
              <p className={"author"}>{"Young D. Kwon, Jagmohan Chauhan, Abhishek Kumar, Pan Hui, Cecilia Mascolo"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2110.13290"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109GLOBECOM46510.2021.9685498.png"} alt={"Cover image for Context-Aware Augmented Reality with 5G Edge"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Context-Aware Augmented Reality with 5G Edge"}</h4>
              <p className={"conference"}>{"GLOBECOM | December 2021"}</p>
              <p className={"author"}>{"Jacky Cao, Xiaoli Liu, Xiang Su, Sasu Tarkoma, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/GLOBECOM46510.2021.9685498"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109MSN53354.2021.00042.png"} alt={"Cover image for Evaluating Multimedia Protocols on 5G Edge for Mobile Augmented Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Evaluating Multimedia Protocols on 5G Edge for Mobile Augmented Reality"}</h4>
              <p className={"conference"}>{"MSN | December 2021"}</p>
              <p className={"author"}>{"Jacky Cao, Xiang Su, Benjamin Finley, Antti Pauanne, Mostafa Ammar, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/MSN53354.2021.00042"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2111.13486.png"} alt={"Cover image for When Creators Meet the Metaverse: A Survey on Computational Arts"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"When Creators Meet the Metaverse: A Survey on Computational Arts"}</h4>
              <p className={"conference"}>{"Technical Report | November 2021"}</p>
              <p className={"author"}>{"Lik-Hang Lee, Zijun Lin, Rui Hu, Zhengya Gong, Abhishek Kumar, Tangyao Li, Sijia Li, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2111.13486"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109JPROC.2021.3119950.png"} alt={"Cover image for Edge Intelligence: Empowering Intelligence to the Edge of Network"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Edge Intelligence: Empowering Intelligence to the Edge of Network"}</h4>
              <p className={"conference"}>{"IEEE | November 2021"}</p>
              <p className={"author"}>{"Dianlei Xu, Tong Li, Yong Li, Xiang Su, Sasu Tarkoma, Tao Jiang, Jon Crowcroft, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/JPROC.2021.3119950"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TNSE.2021.3131194.png"} alt={"Cover image for Finding Spatiotemporal Patterns of Mobile Application Usage"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Finding Spatiotemporal Patterns of Mobile Application Usage"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Network Science and Engineering | November 2021"}</p>
              <p className={"author"}>{"Tong Li, Yong Li, Tong Xia, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TNSE.2021.3131194"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453476087.png"} alt={"Cover image for This Website Uses Nudging: MTurk Workers' Behaviour on Cookie Consent Notices"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"This Website Uses Nudging: MTurk Workers' Behaviour on Cookie Consent Notices"}</h4>
              <p className={"conference"}>{"CSCW | November 2021"}</p>
              <p className={"author"}>{"Carlos Bermejo Fernandez, Dimitris Chatzopoulos, Dimitris Papadopoulos, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3476087"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2102.00423.png"} alt={"Cover image for Characterizing Student Engagement Moods for Dropout Prediction in Question Pool Websites"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Characterizing Student Engagement Moods for Dropout Prediction in Question Pool Websites"}</h4>
              <p className={"conference"}>{"CSCW | November 2021"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Xiaojuan Ma, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2102.00423"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453487351.3488328.png"} alt={"Cover image for IAN: Interpretable Attention Network for Churn Prediction in LBSNs"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"IAN: Interpretable Attention Network for Churn Prediction in LBSNs"}</h4>
              <p className={"conference"}>{"ASONAM | November 2021"}</p>
              <p className={"author"}>{"Liang-Yu Chen, Yutong Chen, Young D. Kwon, Youwen Kang, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3487351.3488328"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453487351.3488353.png"} alt={"Cover image for Interpretable Business Survival Prediction"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Interpretable Business Survival Prediction"}</h4>
              <p className={"conference"}>{"ASONAM | November 2021"}</p>
              <p className={"author"}>{"Anish Krishna Vallapuram, Nikhil Nanda, Young Dae Kwon, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3487351.3488353"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453474085.3475413.png"} alt={"Cover image for A2W: Context-Aware Recommendation System for Mobile Augmented Reality Web Browser"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A2W: Context-Aware Recommendation System for Mobile Augmented Reality Web Browser"}</h4>
              <p className={"conference"}>{"ACM Multimedia 2021 | October 2021"}</p>
              <p className={"author"}>{"Kit-Yung Lam*, Lik-Hang Lee*, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3474085.3475413"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453474085.3475507.png"} alt={"Cover image for Theophany: Multimodal Speech Augmentation in Instantaneous Privacy Channels"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Theophany: Multimodal Speech Augmentation in Instantaneous Privacy Channels"}</h4>
              <p className={"conference"}>{"ACM Multimedia 2021 | October 2021"}</p>
              <p className={"author"}>{"Abhishek Kumar, Tristan Braud, Lik-Hang Lee, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3474085.3475507"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453474085.3475552.png"} alt={"Cover image for Seeing is Believing? Effects of Visualization on Smart Device Privacy Perceptions"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Seeing is Believing? Effects of Visualization on Smart Device Privacy Perceptions"}</h4>
              <p className={"conference"}>{"ACM Multimedia 2021 | October 2021"}</p>
              <p className={"author"}>{"Carlos Bermejo Fernandez, Lik-Hang Lee, Petteri Nurmi, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3474085.3475552"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453462244.3479885.png"} alt={"Cover image for PARA: Privacy Management and Control in Emerging IoT Ecosystems using Augmented Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"PARA: Privacy Management and Control in Emerging IoT Ecosystems using Augmented Reality"}</h4>
              <p className={"conference"}>{"ICMI | October 2021"}</p>
              <p className={"author"}>{"Carlos Bermejo Fernandez, Petteri Nurmi, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3462244.3479885"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2110.05352.png"} alt={"Cover image for All One Needs to Know about Metaverse: A Complete Survey on Technological Singularity"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"All One Needs to Know about Metaverse: A Complete Survey on Technological Singularity"}</h4>
              <p className={"conference"}>{"Technical Report | October 2021"}</p>
              <p className={"author"}>{"Lik-Hang Lee, Tristan Braud, Pengyuan Zhou, Lin Wang, Dianlei Xu, Zijun Lin, Abhishek Kumar, Carlos Bermejo Fernandez, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2110.05352"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453463525.png"} alt={"Cover image for Who Will Survive and Revive Undergoing the Epidemic: Analyses about POI Visit Behaviour"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Who Will Survive and Revive Undergoing the Epidemic: Analyses about POI Visit Behaviour"}</h4>
              <p className={"conference"}>{"IMWUT/UbiComp | September 2021"}</p>
              <p className={"author"}>{"Zhenyu Han, Haohan Fu, Fengli Xu, Zhen Tu, Yang Yu, Pan Hui, Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3463525"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453432229.png"} alt={"Cover image for HealthWalks: Sensing Fine-grained Individual Health Condition via Mobility Data"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"HealthWalks: Sensing Fine-grained Individual Health Condition via Mobility Data"}</h4>
              <p className={"conference"}>{"IMWUT/UbiComp | September 2021"}</p>
              <p className={"author"}>{"Zongyu Lin, Shiqing Lyu, Hancheng Gao, Fengli Xu, Yuqiong Wei, Pan Hui, Hanan Samet, Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3432229"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453432205.png"} alt={"Cover image for Aquilis: Using Contextual Integrity for Privacy Protection on Mobile Devices"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Aquilis: Using Contextual Integrity for Privacy Protection on Mobile Devices"}</h4>
              <p className={"conference"}>{"IMWUT/UbiComp | September 2021"}</p>
              <p className={"author"}>{"Abhishek Kumar, Tristan Braud, Young Dae Kwon, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3432205"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453432203.png"} alt={"Cover image for ContAuth: Continual Learning Framework for Behavioral-based User Authentication"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"ContAuth: Continual Learning Framework for Behavioral-based User Authentication"}</h4>
              <p className={"conference"}>{"IMWUT/UbiComp | September 2021"}</p>
              <p className={"author"}>{"Jagmohan Chauhan, Young Kwon, Pan Hui, Cecilia Mascolo"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3432203"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453448078.png"} alt={"Cover image for Passive Health Monitoring using Large Scale Mobility Data"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Passive Health Monitoring using Large Scale Mobility Data"}</h4>
              <p className={"conference"}>{"IMWUT/UbiComp | September 2021"}</p>
              <p className={"author"}>{"Yunke Zhang, Fengli Xu, Tong Li, Vassilis Kostakos, Pan Hui, Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3448078"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109MC.2021.3112850.png"} alt={"Cover image for Collaboration Stability: Quantifying the Success and Failure of Opportunistic Collaboration"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Collaboration Stability: Quantifying the Success and Failure of Opportunistic Collaboration"}</h4>
              <p className={"conference"}>{"IEEE Computer Magazine | September 2021"}</p>
              <p className={"author"}>{"Huber Flores, Agustin Zuniga, Leonardo Tonetto, Tristan Braud, Pan Hui, Yong Li, Sasu Tarkoma, Mostafa Ammar, Petteri Nurmi"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/MC.2021.3112850"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2101.04825.png"} alt={"Cover image for Towards Mobile Distributed Ledgers"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards Mobile Distributed Ledgers"}</h4>
              <p className={"conference"}>{"IEEE Internet of Things Journal | September 2021"}</p>
              <p className={"author"}>{"Dimitris Chatzopoulos, Anurag Jain, Sujit Gujar, Boi Faltings, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2101.04825"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453447548.3467181.png"} alt={"Cover image for Hierarchical Reinforcement Learning on Scarce Medical Resource Allocation with Imperfect Information"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Hierarchical Reinforcement Learning on Scarce Medical Resource Allocation with Imperfect Information"}</h4>
              <p className={"conference"}>{"KDD | August 2021"}</p>
              <p className={"author"}>{"Qianyue Hao, Fengli Xu, Lin Chen, Pan Hui, Yong Li"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3447548.3467181"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TKDE.2021.3110724.png"} alt={"Cover image for Persuade to Click: Context-aware Persuasion Model for Online Textual Advertisement"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Persuade to Click: Context-aware Persuasion Model for Online Textual Advertisement"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Knowledge and Data Engineering | August 2021"}</p>
              <p className={"author"}>{"Yuan Yuan, Fengli Xu, Hancheng Cao, Guozhen Zhang, Pan Hui, Yong Li, Depeng Jin"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TKDE.2021.3110724"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ICDCS51616.2021.00074.png"} alt={"Cover image for CAD3: Edge-facilitated Real-time Collaborative Abnormal Driving Distributed Detection"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"CAD3: Edge-facilitated Real-time Collaborative Abnormal Driving Distributed Detection"}</h4>
              <p className={"conference"}>{"ICDCS | July 2021"}</p>
              <p className={"author"}>{"Ahmad Alhilal, Tristan Braud, Xiang Su, Luay Al Asadi, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ICDCS51616.2021.00074"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TMC.2021.3098664.png"} alt={"Cover image for Understanding the Long-term Evolution of Mobile App Usage"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Understanding the Long-term Evolution of Mobile App Usage"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Mobile Computing | July 2021"}</p>
              <p className={"author"}>{"Tong Li, Yali Fan, Yong Li, Sasu Tarkoma, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TMC.2021.3098664"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ACCESS.2021.3094828.png"} alt={"Cover image for A Roadmap Towards a Unified Space Communication Architecture"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A Roadmap Towards a Unified Space Communication Architecture"}</h4>
              <p className={"conference"}>{"IEEE Access | June 2021"}</p>
              <p className={"author"}>{"Ahmad Alhilal, Tristan Braud, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ACCESS.2021.3094828"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453430895.3460126.png"} alt={"Cover image for Student Barriers to Active Learning in Synchronous Online Classes: Characterization, Reflections, and Suggestions"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Student Barriers to Active Learning in Synchronous Online Classes: Characterization, Reflections, and Suggestions"}</h4>
              <p className={"conference"}>{"L@S2021 | June 2021"}</p>
              <p className={"author"}>{"Reza Hadi Mogavi, Yankun Zhao, Ehsan Ul Haq, Pan Hui, Xiaojuan Ma"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3430895.3460126"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453457141.png"} alt={"Cover image for Exploring Button Designs for Mid-air Interaction in Virtual Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Exploring Button Designs for Mid-air Interaction in Virtual Reality"}</h4>
              <p className={"conference"}>{"PACM HCI/EICS | June 2021"}</p>
              <p className={"author"}>{"Carlos Bermejo Fernandez, Lik-Hang Lee, Paul Chojecki, David Przewozny, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3457141"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453457146.png"} alt={"Cover image for Press-n-Paste: Copy-and-Paste Operations with Pressure-sensitive Caret Navigation for Miniaturized Surface in Mobile Augmented Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Press-n-Paste: Copy-and-Paste Operations with Pressure-sensitive Caret Navigation for Miniaturized Surface in Mobile Augmented Reality"}</h4>
              <p className={"conference"}>{"PACM HCI/EICS | June 2021"}</p>
              <p className={"author"}>{"Lik-Hang Lee, Yiming Zhu, Yui-Pan Yau, Pan Hui, Susanna Pirttikangas"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3457146"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109TKDE.2021.3091503.png"} alt={"Cover image for DeepPick: A Deep Learning Approach to Unveil Outstanding Users Ranking with Public Attainable Features"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"DeepPick: A Deep Learning Approach to Unveil Outstanding Users Ranking with Public Attainable Features"}</h4>
              <p className={"conference"}>{"IEEE Transactions on Knowledge and Data Engineering | May 2021"}</p>
              <p className={"author"}>{"Wanda Li, Zhiwei Xu, Yi Sun, Qingyuan Gong, Yang Chen, Aaron Yi Ding, Xin Wang, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/TKDE.2021.3091503"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453467963.png"} alt={"Cover image for Towards Augmented Reality Driven Human-City Interaction: Current Research on Mobile Headsets and Future Challenges"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Towards Augmented Reality Driven Human-City Interaction: Current Research on Mobile Headsets and Future Challenges"}</h4>
              <p className={"conference"}>{"ACM Computing Survey | May 2021"}</p>
              <p className={"author"}>{"Lik-Hang Lee, Tristan Braud, Simo Hosio, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3467963"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453465396.png"} alt={"Cover image for A survey on Haptic Technologies for Mobile Augmented Reality"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"A survey on Haptic Technologies for Mobile Augmented Reality"}</h4>
              <p className={"conference"}>{"ACM Computing Survey | May 2021"}</p>
              <p className={"author"}>{"Carlos Bermejo Fernandez, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3465396"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1371journal.pone.0251550.png"} alt={"Cover image for Deep Reinforcement Learning Approaches for Global Public Health Strategies for COVID-19 Pandemic"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Deep Reinforcement Learning Approaches for Global Public Health Strategies for COVID-19 Pandemic"}</h4>
              <p className={"conference"}>{"PLOS ONE | May 2021"}</p>
              <p className={"author"}>{"Gloria Hyunjung Kwak, Lowell Ling, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1371/journal.pone.0251550"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453442381.3449829.png"} alt={"Cover image for DeepVista: 16K Panoramic Cinema on Your Mobile Device"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"DeepVista: 16K Panoramic Cinema on Your Mobile Device"}</h4>
              <p className={"conference"}>{"WWW 2021 | April 2021"}</p>
              <p className={"author"}>{"Wenxiao Zhang, Feng Qian, Bo Han, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3442381.3449829"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109JIOT.2021.3073864.png"} alt={"Cover image for The Impact of Covid-19 on Smartphone Usage"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"The Impact of Covid-19 on Smartphone Usage"}</h4>
              <p className={"conference"}>{"IEEE Internet of Things Journal | April 2021"}</p>
              <p className={"author"}>{"Tong Li, Mingyang Zhang, Yong Li, Eemil Lagerspetz, Sasu Tarkoma, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/JIOT.2021.3073864"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453457950.png"} alt={"Cover image for Emerging ExG-based NUI Inputs in Extended Realities: A Bottom-Up Survey."} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Emerging ExG-based NUI Inputs in Extended Realities: A Bottom-Up Survey."}</h4>
              <p className={"conference"}>{"ACM Transactions on Interactive Intelligent Systems | March 2021"}</p>
              <p className={"author"}>{"Kirill A. Shatilov, Dimitris Chatzopoulos, Lik-Hang Lee, and Pan Hui."}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3457950"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.11453451394.png"} alt={"Cover image for 3DGCN: Dynamic 3D Graph Convolutional Network for Citywide Crowd Flow Prediction"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"3DGCN: Dynamic 3D Graph Convolutional Network for Citywide Crowd Flow Prediction"}</h4>
              <p className={"conference"}>{"ACM Transactions on Knowledge Discovery from Data (TKDD) | February 2021"}</p>
              <p className={"author"}>{"Tong Xia, Junjie Lin, Feng Yong, Jie Feng, Pan Hui, Depeng Jin"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1145/3451394"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109MIC.2021.3059189.png"} alt={"Cover image for Intelligent and Scalable Air Quality Monitoring with 5G Edge"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Intelligent and Scalable Air Quality Monitoring with 5G Edge"}</h4>
              <p className={"conference"}>{"IEEE Internet Computing | February 2021"}</p>
              <p className={"author"}>{"Xiang Su, Xiaoli Liu, Naser Hossein Motlagh, Jacky Cao, Peifeng Su, Petri Pellikka, Yongchun Liu, Tuukka Petäjä, Markku Kulmala, Pan Hui, Sasu Tarkoma"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/MIC.2021.3059189"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.23919JSC.2021.0004.png"} alt={"Cover image for DeepPredict: A Zone Preference Prediction System for Online Lodging Platforms"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"DeepPredict: A Zone Preference Prediction System for Online Lodging Platforms"}</h4>
              <p className={"conference"}>{"Journal of Social Computing | 2021"}</p>
              <p className={"author"}>{"Yihan Ma, Hua Sun, Yang Chen, Jiayun Zhang, Yang Xu, Xin Wang, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.23919/JSC.2021.0004"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.48550arXiv.2101.06443.png"} alt={"Cover image for Predicting Hyperkalemia in the ICU and Evaluation of Generalizability and Interpretability"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Predicting Hyperkalemia in the ICU and Evaluation of Generalizability and Interpretability"}</h4>
              <p className={"conference"}>{"AAAI 2021 Workshop on Trustworthy AI for Healthcare | February 2021"}</p>
              <p className={"author"}>{"Hyunjung Kwak, Christina Chen, Lowell Ling, Erina Ghosh, Leo Celi, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.48550/arXiv.2101.06443"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1109ACCESS.2021.3050929.png"} alt={"Cover image for Sketching an AI Marketplace: Tech, Economic, and Regulatory Aspects"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Sketching an AI Marketplace: Tech, Economic, and Regulatory Aspects"}</h4>
              <p className={"conference"}>{"IEEE Access | January 2021"}</p>
              <p className={"author"}>{"Abhishek Kumar, Benjamin Finley, Tristan Braud, Sasu Tarkoma, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1109/ACCESS.2021.3050929"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
        <div className={"row"}>
          <div className={"row"}>
            <div className={"col-md-4"}>
              <img src={"/images/Publication_png/10.1049blc2.12001.png"} alt={"Cover image for Decentralizing Indexing and Bootstrapping for Online Applications"} onError={useFallbackPhoto} />
            </div>
            <div className={"col-md-8"}>
              <h4>{"Decentralizing Indexing and Bootstrapping for Online Applications"}</h4>
              <p className={"conference"}>{"IET Blockchain Journal (Invited Paper) | 2021"}</p>
              <p className={"author"}>{"Pierre Schutz, Stanislas Gal, Dimitris Chatzopoulos, Pan Hui"}</p>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-file-pdf"}></i>
                {" PDF"}
              </a>
              <a href={""} className={"publication-btn"} target={"_blank"}>
                <i className={"fa fa-video-camera"}></i>
                {" Video"}
              </a>
              <a href={"https://doi.org/10.1049/blc2.12001"} className={"publication-btn"} target={"_blank"}>{"DOI"}</a>
            </div>
          </div>
        </div>
      </section>
    </>
  );
}

function FloatingContactMenu(): JSX.Element {
  const [isOpen, setIsOpen] = useState(false);
  useEffect(() => {
    if (!isOpen) return undefined;
    const closeMenu = (event: MouseEvent | TouchEvent) => {
      const target = event.target;
      if (target instanceof Element && target.closest('.menu')) return;
      setIsOpen(false);
    };
    window.addEventListener('click', closeMenu);
    window.addEventListener('touchstart', closeMenu);
    return () => {
      window.removeEventListener('click', closeMenu);
      window.removeEventListener('touchstart', closeMenu);
    };
  }, [isOpen]);

  return (
    <div className='wrapMenu'>
      <div className={isOpen ? 'menu menu--bottom-right menu-closed active' : 'menu menu--bottom-right menu-closed'} id='menu_bottom_right'>
        <a className='menu__btn' href='#menu_bottom_right' onClick={(event) => { event.preventDefault(); setIsOpen((open) => !open); }}><i className='fa fa-info'></i></a>
        <ul className='menu__list'>
          <li><a style={{ color: '#fff' }} href='mailto:mc2@hkust-gz.edu.cn'><i className='fa-solid fa-envelope'></i></a></li>
          <li><a style={{ color: '#fff' }} href='#wechat'><i className='fa-brands fa-weixin'></i></a></li>
          <li><a style={{ color: '#fff' }} href='https://medium.com/@mc2.hkust.hkustgz'><i className='fa-brands fa-medium'></i></a></li>
          <li><a style={{ color: '#fff' }} href='https://www.youtube.com/@MC2HKUSTGZCWB'><i className='fa-brands fa-youtube'></i></a></li>
        </ul>
      </div>
    </div>
  );
}

export default function PublicationPage({ locale = 'en' }: PublicationPageProps): JSX.Element {
  const pageRef = useRef<HTMLElement>(null);
  const [activeYear, setActiveYear] = useState('2026');
  const isZh = locale === 'zh-CN';

  useBodyPublicationClass();
  usePublicationScrollSpy(pageRef, setActiveYear);

  useEffect(() => {
    if (pageRef.current) decoratePanHuiNames(pageRef.current);
  }, [locale]);

  const selectYear = (year: string) => {
    setActiveYear(year);
    const section = pageRef.current?.querySelector<HTMLElement>(`#yr${year}`);
    const scrollContainer = pageRef.current?.querySelector<HTMLElement>('.publication-section');
    if (!section || !scrollContainer) return;
    const containerRect = scrollContainer.getBoundingClientRect();
    const sectionTop = section.getBoundingClientRect().top - containerRect.top + scrollContainer.scrollTop;
    const previousScrollBehavior = scrollContainer.style.getPropertyValue('scroll-behavior');
    const previousScrollBehaviorPriority = scrollContainer.style.getPropertyPriority('scroll-behavior');
    scrollContainer.style.setProperty('scroll-behavior', 'auto', 'important');
    scrollContainer.scrollTop = Math.max(0, sectionTop - 90);
    window.requestAnimationFrame(() => {
      if (previousScrollBehavior) scrollContainer.style.setProperty('scroll-behavior', previousScrollBehavior, previousScrollBehaviorPriority);
      else scrollContainer.style.removeProperty('scroll-behavior');
    });
    window.history.replaceState(null, '', `#yr${year}`);
  };

  return (
    <main className='publication-page-shell' aria-label={isZh ? '论文' : 'Publication'} ref={pageRef}>
      <PublicationNavBar locale={locale} />
      <div className='publication-section'>
        <div className='container'>
          <div className='row justify-content-center'>
            <div className='col-12 text-center' style={{ paddingTop: '100px' }}>
              <h3>{isZh ? '论文' : 'Publication'}</h3>
            </div>
            <YearNav activeYear={activeYear} onSelectYear={selectYear} />
            <div className='col-12 mt-3'>
              <RecentPublicationSections />
              <StaticPublicationSections />
            </div>
          </div>
        </div>
      </div>
      <FloatingContactMenu />
    </main>
  );
}
