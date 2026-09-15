/* Video requests are deferred until an explicit visitor action. */
(() => {
  document.querySelectorAll('[data-learning-video]').forEach((box) => {
    const button = box.querySelector('button');
    const stage = box.querySelector('.official-video-stage');
    const id = box.dataset.learningVideo;
    if (!button || !stage || !/^[A-Za-z0-9_-]{11}$/.test(id)) return;
    button.hidden = false;
    button.addEventListener('click', () => {
      const frame = document.createElement('iframe');
      frame.src = `https://www.youtube-nocookie.com/embed/${id}?autoplay=1&rel=0`;
      frame.title = '와와학습코칭센터 공식 채널 원장 인터뷰와 학원 후기';
      frame.allow = 'autoplay; encrypted-media; picture-in-picture; fullscreen';
      frame.allowFullscreen = true;
      frame.referrerPolicy = 'strict-origin-when-cross-origin';
      stage.replaceChildren(frame);
      frame.focus();
    }, { once: true });
  });
})();
