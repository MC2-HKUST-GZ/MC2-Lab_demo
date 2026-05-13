import React, { type RefObject } from 'react';

interface HomeSocialMenuProps {
  isOpen: boolean;
  menuRef: RefObject<HTMLDivElement>;
  onToggle: () => void;
}

export default function HomeSocialMenu({ isOpen, menuRef, onToggle }: HomeSocialMenuProps): JSX.Element {
  return (
    <div className='wrapMenu' ref={menuRef}>
      <div className={isOpen ? 'menu menu--bottom-right active' : 'menu menu--bottom-right menu-closed'} id='menu_bottom_right'>
        <a
          className='menu__btn'
          href='#contact-menu'
          aria-expanded={isOpen}
          aria-label='Contact links'
          onClick={(event) => {
            event.preventDefault();
            onToggle();
          }}
        >
          <i className='fa fa-info' aria-hidden='true'></i>
        </a>
        <ul className='menu__list'>
          <li><a href='mailto:mc2@hkust-gz.edu.cn' aria-label='Email MC2'><i className='fa-solid fa-envelope' aria-hidden='true'></i></a></li>
          <li><a href='#' aria-label='WeChat'><i className='fa-brands fa-weixin' aria-hidden='true'></i></a></li>
          <li><a href='https://medium.com/@mc2.hkust.hkustgz' aria-label='Medium'><i className='fa-brands fa-medium' aria-hidden='true'></i></a></li>
          <li><a href='https://www.youtube.com/@MC2HKUSTGZCWB' aria-label='YouTube'><i className='fa-brands fa-youtube' aria-hidden='true'></i></a></li>
        </ul>
      </div>
    </div>
  );
}

