import React, { useEffect, useState } from 'react';

export default function FloatingContactMenu(): JSX.Element {
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
        <a className='menu__btn' href='#menu_bottom_right' aria-expanded={isOpen} aria-label='Contact links' onClick={(event) => { event.preventDefault(); setIsOpen((open) => !open); }}>
          <i className='fa fa-info' aria-hidden='true'></i>
        </a>
        <ul className='menu__list'>
          <li><a style={{ color: '#fff' }} href='mailto:mc2@hkust-gz.edu.cn' aria-label='Email MC2'><i className='fa-solid fa-envelope' aria-hidden='true'></i></a></li>
          <li><a style={{ color: '#fff' }} href='#wechat' aria-label='WeChat'><i className='fa-brands fa-weixin' aria-hidden='true'></i></a></li>
          <li><a style={{ color: '#fff' }} href='https://medium.com/@mc2.hkust.hkustgz' aria-label='Medium'><i className='fa-brands fa-medium' aria-hidden='true'></i></a></li>
          <li><a style={{ color: '#fff' }} href='https://www.youtube.com/@MC2HKUSTGZCWB' aria-label='YouTube'><i className='fa-brands fa-youtube' aria-hidden='true'></i></a></li>
        </ul>
      </div>
    </div>
  );
}
