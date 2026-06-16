export const project = {
  name: 'project',
  title: 'Project',
  type: 'document',
  fields: [
    { name: 'title', title: 'Title', type: 'string' },
    { 
      name: 'slug', 
      title: 'Slug (URL ID)', 
      type: 'slug', 
      options: { source: 'title', maxLength: 96 } 
    },
    {
      name: 'category',
      title: 'Categories',
      type: 'array',
      of: [
        {
          type: 'string',
          options: {
            list: [
              { title: 'Structured Cabling', value: 'Structured Cabling' },
              { title: 'A/V Systems', value: 'A/V Systems' },
              { title: 'Integrated Systems', value: 'Integrated Systems' },
              { title: 'Security Systems', value: 'Security Systems' }
            ]
          }
        }
      ]
    },
    { name: 'image', title: 'Main Image', type: 'image', options: { hotspot: true } },
    { name: 'description', title: 'Description', type: 'text' },
    { name: 'client', title: 'Client', type: 'string' },
    { name: 'duration', title: 'Duration', type: 'string' },
    { name: 'challenge', title: 'The Challenge', type: 'text' },
    { name: 'solution', title: 'The Solution', type: 'text' },
    { 
      name: 'gearList', 
      title: 'Technology Deployed', 
      type: 'array', 
      of: [{type: 'string'}] 
    },
    {
      name: 'stats',
      title: 'Stats',
      type: 'array',
      of: [
        {
          type: 'object',
          fields: [
            { name: 'label', title: 'Label', type: 'string' },
            { name: 'value', title: 'Value', type: 'string' }
          ]
        }
      ]
    }
  ]
}