import React from 'react';
import { Link } from 'react-router-dom';
import type { NewsArticleBlock, NewsArticleContent as NewsArticleContentType } from '../../types/legacyPages';
import InlineMc2Text from './InlineMc2Text';

interface NewsArticleContentProps {
  article: NewsArticleContentType;
}

function ArticleBlock({ block }: { block: NewsArticleBlock }): JSX.Element {
  if (block.type === 'heading') return <h4><InlineMc2Text text={block.text} /></h4>;
  if (block.type === 'list') {
    return (
      <ul className='article-list'>
        {block.items.map((item) => (
          <li key={item}>{item}</li>
        ))}
      </ul>
    );
  }
  return <p><InlineMc2Text text={block.text} /></p>;
}

export default function NewsArticleContent({ article }: NewsArticleContentProps): JSX.Element {
  return (
    <>
      <div className='news-section'>
        <div className='container'>
          <div className='row justify-content-center'>
            <div className='col-12 text-center' style={{ paddingTop: '100px', color: '#fff' }}>
              <h3>{article.pageHeading}</h3>
            </div>

            <div className='col-lg-10 mt-4 mb-5'>
              <article className='news-article-card'>
                <h2 className='article-title'><InlineMc2Text text={article.title} /></h2>
                <p className='article-meta'><InlineMc2Text text={article.meta} /></p>
                <img className='article-cover' src={article.image} alt={article.imageAlt} />
                {article.blocks.map((block, index) => (
                  <ArticleBlock block={block} key={`${block.type}-${index}`} />
                ))}
                <div className='article-links'>
                  {article.links.map((link) => (
                    <a href={link.href} key={link.href} target='_blank' rel='noopener noreferrer'>{link.label}</a>
                  ))}
                </div>
                <Link className='back-news-link' to={article.backHref}>{article.backLabel}</Link>
              </article>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
