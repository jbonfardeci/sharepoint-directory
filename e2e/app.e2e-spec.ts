import { SharepointDirectoryPage } from './app.po';

describe('sharepoint-directory App', function() {
  let page: SharepointDirectoryPage;

  beforeEach(() => {
    page = new SharepointDirectoryPage();
  });

  it('should display message saying app works', () => {
    page.navigateTo();
    expect(page.getParagraphText()).toEqual('app works!');
  });
});
