import { join } from 'path';

export const templatesPath = {
  confirmation: join(process.cwd(), 'dist/assets/templates/confirmationTemplate.html'),
  passwordReset: join(process.cwd(), 'dist/assets/templates/passReset.html'),
};
