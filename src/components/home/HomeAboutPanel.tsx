import React, { type CSSProperties } from 'react';
import { homeAboutCopy, homeResearchFocus } from '../../data/homePageContent';
import HomePanelCloseButton from './HomePanelCloseButton';

interface HomeAboutPanelProps {
  onClosePanel: () => void;
}

const aboutSectionStyle: CSSProperties = { overflow: 'auto' };
const aboutSectionCenterStyle: CSSProperties = { transform: 'none', top: '10%' };
const aboutHeadingStyle: CSSProperties = { fontSize: 'min(10vw,10vh)', fontFamily: 'Open Sans', fontWeight: 400 };
const justifiedTextStyle: CSSProperties = { textAlign: 'justify' };
const justifiedInterWordStyle: CSSProperties = { textAlign: 'justify', textJustify: 'inter-word' };
const centeredRowStyle: CSSProperties = { display: 'flex', justifyContent: 'center' };

export default function HomeAboutPanel({ onClosePanel }: HomeAboutPanelProps): JSX.Element {
  return (
    <section className='about-section' style={aboutSectionStyle}>
      <div className='section-center' style={aboutSectionCenterStyle}>
        <div className='container'>
          <div className='row justify-content-center'>
            <div className='col-12 text-center'>
              <p style={aboutHeadingStyle}>About Us</p>
            </div>
            <div className='col-lg-12 mt-4' style={justifiedTextStyle}>
              <p>{homeAboutCopy.intro}</p>
            </div>
            <div className='col-lg-12 mt-4' style={justifiedInterWordStyle}>
              <p>
                {homeAboutCopy.leaderPrefix}
                <a href={homeAboutCopy.leaderHref} target='_blank' rel='noopener noreferrer'>{homeAboutCopy.leaderName}</a>
                {homeAboutCopy.leaderSuffix}
              </p>
            </div>
            <div className='col-lg-12 mt-4' style={justifiedInterWordStyle}>
              <ul>
                {homeResearchFocus.map((item) => (
                  <li key={item.title}>• <strong>{item.title}</strong>: {item.body}</li>
                ))}
              </ul>
            </div>
            <div className='col-lg-12 mt-4' style={centeredRowStyle}>
              <p>{homeAboutCopy.closing}</p>
            </div>
            <div className='col-lg-12 mt-4' style={centeredRowStyle}>
              <HomePanelCloseButton className='close_btn about-close_btn' onClose={onClosePanel} />
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}
