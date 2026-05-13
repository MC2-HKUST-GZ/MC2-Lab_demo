import React from 'react';
import FloatingContactMenu from '../FloatingContactMenu';
import type { ProjectItem, ProjectPageContent } from '../../types/legacyPages';

interface ProjectContentProps {
  content: ProjectPageContent;
}

function ProjectCard({ project }: { project: ProjectItem }): JSX.Element {
  const body = (
    <div className={project.href ? 'row project-spotlight-card' : 'row'} style={{ marginBottom: '2em' }}>
      <div className='col-md-5' style={{ display: 'flex', justifyContent: project.id === 'project-5' ? 'start' : 'center', alignItems: 'flex-start', paddingTop: '30px' }}>
        <img src={project.image} alt={project.imageAlt} className={project.href ? 'project-spotlight-image' : undefined} style={project.href ? undefined : { maxWidth: '100%' }} />
      </div>
      <div className='col-md-7' style={{ display: 'flex', flexDirection: 'column', gap: project.id === 'project-1' ? '.25em' : '.5em' }}>
        <h4 style={{ fontSize: '2.0rem' }}>{project.title}</h4>
        <p className='people'>{project.people}</p>
        <div>
          {project.tags.map((tag) => (
            <div className='publication-btn' style={{ marginBottom: '.5em' }} key={tag}>{tag}</div>
          ))}
        </div>
        <div>
          {project.descriptions.map((description) => (
            <p className='description' style={{ textAlign: 'justify' }} key={description}>{description}</p>
          ))}
        </div>
      </div>
    </div>
  );

  if (!project.href) return body;

  return (
    <a className='project-spotlight-link' href={project.href} target='_blank' rel='noopener noreferrer' aria-label='Open Surreality project website'>
      {body}
    </a>
  );
}

export default function ProjectContent({ content }: ProjectContentProps): JSX.Element {
  return (
    <>
      <div className='publication-section'>
        <div className='container'>
          <div className='row justify-content-center'>
            <div className='col-12 text-center' style={{ paddingTop: '100px' }}>
              <h3>{content.pageHeading}</h3>
            </div>

            <div className='col-12 mt-3'>
              <section id='yr2024'>
                {content.projects.map((project) => (
                  <ProjectCard key={project.id} project={project} />
                ))}
              </section>
            </div>
          </div>
        </div>
      </div>
      <FloatingContactMenu />
    </>
  );
}
