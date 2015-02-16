<?php

namespace Tag\AdminBundle\Controller;

/**
 * Comty controller.
 * @Route("/dashboard/community")
 */
class ComtyController extends Controller
{

    private $labels = [0 => 'Client', 1 => 'Brand', 2 => 'Prospect', 3 => 'Clone', 4 => 'Payed Clone', 5 => 'Agency', 6 => 'Marketplace', 7 => 'Template', 8 => 'Partner'];

    /** @var  Comty */
    private $comty;

    public function setContainer(ContainerInterface $container = null){
        parent::setContainer($container);

        $this->comty = $this->get('tag.community.manager')->getCurrentCommunity();
    }

    /**
     * Lists all Comty entities.
     * @Route("/", name="dashboard_community")
     * @Method("GET")
     * @Template()
     */
    public function indexAction()
    {
        $em = $this->getDoctrine()->getManager();
        $entities = $this->get('tagg.com_repository')->findNonCloneCommunities();
        $clones = $this->get('tagg.com_repository')->findCloneCommunities();

        $community = new Comty();
        $create_form = $this->createForm(new ComtyType('create', $community), $community);

        $clone_form = $this->createForm(new CreateCloneType());
        $agency_form = $this->createForm(new CreateCloneType(true));

        return [
            'entities' => $entities,
            'basehost' => $this->container->getParameter('base_host'),
            'labels' => $this->labels,
            'create_form' => $create_form->createView(),
            'clone_form' => $clone_form->createView(),
            'agency_form' => $agency_form->createView(),
            'clones'      => $clones
        ];
    }

    /**
     * Creates a new Comty entity.
     *
     * @Route("/", name="dashboard_community_create")
     * @Method("POST")
     * @Template("TagAdminBundle:Comty:new.html.twig")
     */
    public function createAction(Request $request)
    {
        $entity  = new Comty();
        $form = $this->createForm(new ComtyType('create', $entity), $entity);
        $form->handleRequest($request);
        if ($form->isValid()) {
            $em = $this->getDoctrine()->getManager();
            $entity->setPlan('free');
            $entity->setSlug(Urlizer::transliterate($entity->getSlug()));
            $em->persist($entity);
            $role = new Role();
            $em->flush();
            $role_name = 'ROLE_COMMUNITY_'.$entity->getId();
            $role->setName($role_name);
            $role->setRole($role_name);
            $em->persist($role);
            $em->flush();

            /* creating default theme for community */
            $communityTheme = $this->get('tagg.comtheme_repository')->createDefTheme($entity);
            $em->persist($communityTheme);


            $my_videos = $this->get('tagg.folder_repository')->createCollection('My Videos', null, $entity);
            $em->persist($my_videos);
            $brand = $this->get('tagg.folder_repository')->createCollection('Partners', null, $entity);
            $em->persist($brand);
            $trash = $this->get('tagg.folder_repository')->createCollection('Trash Can',null,$entity);
            $em->persist($trash);
            $item = new MenuItem();
            $item->setTitle('root');
            $item->setType(0);
            $item->setIcon('');
            $item->setValue('');
            $item->setCommunity($entity);

            $em->persist($item);

            $em->flush();
            $this->get('session')->getFlashBag()->add(
                'notice',
                'Community successfully created'
            );
            return $this->redirect($this->generateUrl('dashboard_community_show', array('id' => $entity->getId())));
        }
        return [
            'entity' => $entity,
            'form'   => $form->createView(),
        ];
    }

    /**
     * @Route("/edit/theme", name="community_theme_edit")
     * @Template()
     */
    public function communityThemeEditAction(Request $request)
    {
        /** @var $comId Comty */
        $comId = $this->get('tag.community.manager')->getCurrentCommunity();

        $em = $this->getDoctrine()->getManager();
        $communityTheme = $comId->getCommunityTheme();

        if (!$communityTheme) {
            $communityTheme = new ComtyTheme();
            $communityTheme->setCommunityId($comId);
        }
        $oldCssFile = $communityTheme->getCssFile();
        if ('default.css' == $oldCssFile) {
            $oldCssFile = null;
        }

        $editForm = $this->createForm(new CommunityThemeType($communityTheme), $communityTheme);
        $editForm->handleRequest($request);

        if ($editForm->isValid()){

            if ($this->isValidImages($communityTheme, $editForm)) {
                $communityTheme = $this->get('tagg.comtheme_repository')->uploadImages($editForm, $communityTheme, $comId, $this->getRequest()->server->get('DOCUMENT_ROOT'));

                $requestData = $request->request->all();
                $newCssFile = $this->get('tagg.comtheme_repository')->generateCss(
                    $oldCssFile,
                    $requestData['community_theme'],
                    $this->getRequest()->server->get('DOCUMENT_ROOT'),
                    $comId->getSlug(),
                    $communityTheme->getBodyBgPath()
                );

                if ($newCssFile) {
                    $communityTheme->setCssFile($newCssFile);
                    $em->persist($communityTheme);
                    $em->flush();

                    $this->get('session')->getFlashBag()->add(
                        'notice',
                        'Changes are successfully saved'
                    );
                } else {
                    $this->get('session')->getFlashBag()->add(
                        'notice',
                        'Can not generate css-file. Theme was not saved.'
                    );
                }
            } else {
                $this->get('session')->getFlashBag()->add(
                    'error',
                    'Please add pictures.'
                );
            }


        }

        return array(
            'theme' => $communityTheme,
            'edit_form' => $editForm->createView()
        );
    }

    /**
     * @Route("/details", name="community_details")
     * @Template()
     */
    public function communityDetailsAction(){
        /** @var Comty $community */
        $community = $this->get('tag.community.manager')->getCurrentCommunity();

        $form = $this->createForm(new CommunityType(), $community);
        if ($this->getRequest()->isMethod('post')){
            $form->submit($this->getRequest());
            if ($form->isValid()){
                $em = $this->getDoctrine()->getManager();
                $em->persist($community);
                $em->flush();
                $this->get('session')->getFlashBag()->add(
                    'notice',
                    'Changes are successfully saved'
                );
                return $this->redirect($this->generateUrl('community_details'));
            }
        }

        return array(
            'form' => $form->createView(),
            'community' => $community
        );
    }

}
